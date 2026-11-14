using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace CaoXing
{
    /// <summary>读取操行汇总表与普通名单表格。</summary>
    public static class XlsxReader
    {
        public static ClassData LoadClass(string path)
        {
            if (!File.Exists(path)) throw new XlsxException("文件不存在：" + path);
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read, false, Encoding.UTF8))
                {
                    return LoadClass(zip, path);
                }
            }
            catch (XlsxException) { throw; }
            catch (Exception ex)
            {
                throw new XlsxException("无法读取该表格，可能不是有效的 .xlsx 文件，或文件正被其他程序占用。\r\n\r\n" + ex.Message, ex);
            }
        }

        private static ClassData LoadClass(ZipArchive zip, string path)
        {
            string sheetPath = XlsxCommon.FindSheetPath(zip);
            string sheetXml = XlsxCommon.ReadEntry(zip, sheetPath);
            if (string.IsNullOrEmpty(sheetXml))
                throw new XlsxException("表格中找不到工作表内容，无法读取。");

            string sharedRaw;
            List<string> shared = XlsxCommon.ReadShared(zip, out sharedRaw);
            XmlDocument sheet = XlsxCommon.LoadXml(sheetXml);
            XmlNamespaceManager ns = XlsxCommon.Ns(sheet);

            List<RowInfo> rows = ReadRows(sheet, ns, shared);
            if (rows.Count == 0) throw new XlsxException("工作表是空的。");

            // 找表头：包含“学号”和“姓名”的行。
            RowInfo header = null;
            int headerRow = 0;
            foreach (RowInfo r in rows)
            {
                if (r.Row > 12) break;
                if (r.HasText("学号") && r.HasText("姓名"))
                {
                    header = r;
                    headerRow = r.Row;
                    break;
                }
            }
            if (header == null)
                throw new XlsxException("没有找到表头。请确认表格第一行附近包含“序号、学号、姓名、模块一……平均分”等列。");

            int seqCol = header.FindColumn("序号");
            int noCol = header.FindColumn("学号");
            int nameCol = header.FindColumn("姓名");
            if (seqCol <= 0) seqCol = 1;
            if (noCol <= 0) noCol = 2;
            if (nameCol <= 0) nameCol = 3;

            int avgCol = header.FindColumn("平均分");
            int moduleCount = 0;
            if (avgCol > 3 && (avgCol - 4) % 2 == 0)
                moduleCount = (avgCol - 4) / 2;
            if (moduleCount <= 0)
            {
                // 没有平均分表头时，按“模块”表头数量推断。
                int count = 0;
                for (int c = 4; c < 60; c++)
                {
                    string text = header.Get(c).Trim();
                    if (text.StartsWith("模块")) count++;
                }
                if (count > 0)
                {
                    moduleCount = count;
                    avgCol = 4 + count * 2;
                }
            }
            if (moduleCount <= 0)
                throw new XlsxException("无法识别模块列。请确认表头中有“模块一”“模块二”……以及“平均分”。");

            // 表头跨行数：A1:A2 这类合并。
            int headerRows = 1;
            foreach (string merge in ReadMergeRefs(sheet, ns))
            {
                int c1, r1, c2, r2;
                if (TryParseRange(merge, out c1, out r1, out c2, out r2))
                {
                    if (c1 == 1 && r1 == headerRow && r2 > r1)
                    {
                        headerRows = Math.Max(headerRows, r2 - r1 + 1);
                    }
                }
            }

            ClassData data = new ClassData();
            data.SourceFile = path;
            data.FromTemplate = true;
            data.DataStartRow = headerRow + headerRows;
            data.AverageColumn = avgCol;
            data.ScoreColumns = new int[moduleCount];
            data.DetailColumns = new int[moduleCount];
            for (int i = 0; i < moduleCount; i++)
            {
                data.ScoreColumns[i] = 4 + i * 2;
                data.DetailColumns[i] = 5 + i * 2;
            }

            for (int i = 0; i < moduleCount; i++)
            {
                string name = header.Get(data.ScoreColumns[i]);
                if (string.IsNullOrWhiteSpace(name)) name = "模块" + Cn.Number(i + 1);
                ModuleDef def = new ModuleDef();
                def.Name = name.Trim();
                def.AspectCount = 5;
                data.Modules.Add(def);
            }

            int maxAspectSeen = 0;
            int[] moduleMaxAspect = new int[moduleCount];
            foreach (RowInfo r in rows)
            {
                if (r.Row < data.DataStartRow) continue;
                string name = r.Get(nameCol).Trim();
                string no = r.Get(noCol).Trim();
                if (name.Length == 0) continue;
                if (IsFooter(name)) continue;
                if (name == "姓名" && no == "学号") continue;

                Student s = new Student();
                s.RowIndex = r.Row;
                s.No = no;
                s.Name = name;
                int seq;
                if (!int.TryParse(r.Get(seqCol).Trim(), out seq)) seq = data.Students.Count + 1;
                s.Seq = seq;

                for (int m = 0; m < moduleCount; m++)
                {
                    ModuleData md = new ModuleData();
                    XmlElement scoreCell = r.GetElement(data.ScoreColumns[m]);
                    XmlElement detailCell = r.GetElement(data.DetailColumns[m]);
                    string scoreText = scoreCell == null ? "" : XlsxCommon.CellText(scoreCell, shared).Trim();
                    decimal score;
                    if (scoreText.Length > 0 && decimal.TryParse(scoreText, NumberStyles.Float, CultureInfo.InvariantCulture, out score))
                    {
                        md.Score = score;
                        md.HasScore = true;
                    }
                    else
                    {
                        md.Score = 0m;
                        md.HasScore = false;
                    }
                    md.OriginalText = detailCell == null ? "" : XlsxCommon.CellText(detailCell, shared);
                    int maxA;
                    md.Records = DetailParser.Parse(md.OriginalText, out maxA);
                    if (maxA > moduleMaxAspect[m]) moduleMaxAspect[m] = maxA;
                    if (maxA > maxAspectSeen) maxAspectSeen = maxA;
                    s.Modules.Add(md);
                }
                data.Students.Add(s);
            }

            // 自定义属性（本程序保存的表格会带这些属性）。
            Dictionary<string, string> props = ReadCustomProps(zip);

            if (data.Students.Count == 0)
                data.Warnings.Add("表格中没有读到学生记录，可以在“导入学生”中导入名单。");

            // 初始分：优先自定义属性，否则取“模块分 - 已解析加减分”的众数。
            decimal initial = 80m;
            bool hasInitialProp = false;
            string initialText;
            if (props.TryGetValue("CX_InitialScore", out initialText))
            {
                decimal p;
                if (decimal.TryParse(initialText.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out p))
                {
                    initial = p;
                    hasInitialProp = true;
                }
            }
            if (!hasInitialProp)
                initial = InferInitialScore(data);

            data.InitialScore = initial;
            for (int si = 0; si < data.Students.Count; si++)
            {
                Student s = data.Students[si];
                for (int m = 0; m < moduleCount; m++)
                {
                    ModuleData md = s.Modules[m];
                    if (!md.HasScore) { md.Score = initial; md.Adjustment = 0m; continue; }
                    decimal adj = Math.Round(md.Score - initial - md.SumDelta(), 3);
                    if (Math.Abs(adj) < 0.0005m) adj = 0m;
                    if (adj != 0m)
                        data.Warnings.Add(string.Format("学生“{0}”的{1}存在无法识别的明细分数，已保留差额 {2} 分。",
                            s.Name, data.Modules[m].Name, Format.Decimal(adj)));
                    md.Adjustment = adj;
                }
            }

            // 方面数量：优先自定义属性，并与实际出现的最大方面号取较大值。
            int[] propAspects = null;
            string aspectsText;
            if (props.TryGetValue("CX_AspectCounts", out aspectsText))
            {
                string[] parts = aspectsText.Split(',');
                if (parts.Length == moduleCount)
                {
                    propAspects = new int[moduleCount];
                    for (int i = 0; i < moduleCount; i++)
                    {
                        int a;
                        if (!int.TryParse(parts[i].Trim(), out a) || a < 1) a = 5;
                        propAspects[i] = a;
                    }
                }
            }
            for (int m = 0; m < moduleCount; m++)
            {
                int a = propAspects != null ? propAspects[m] : 5;
                if (moduleMaxAspect[m] > a) a = moduleMaxAspect[m];
                if (a < 1) a = 1;
                if (a > 30) a = 30;
                data.Modules[m].AspectCount = a;
            }

            string moduleNamesProp;
            if (props.TryGetValue("CX_ModuleNames", out moduleNamesProp))
            {
                string[] names = moduleNamesProp.Split('|');
                if (names.Length == moduleCount)
                {
                    for (int m = 0; m < moduleCount; m++)
                    {
                        if (names[m].Trim().Length > 0) data.Modules[m].Name = names[m].Trim();
                    }
                }
            }

            string classNameProp;
            if (props.TryGetValue("CX_ClassName", out classNameProp) && classNameProp.Trim().Length > 0)
                data.ClassName = classNameProp.Trim();
            else
                data.ClassName = InferClassName(path);

            return data;
        }

        private static bool IsFooter(string name)
        {
            if (name == "备注" || name == "合计" || name == "说明" || name == "总计") return true;
            return false;
        }

        private static decimal InferInitialScore(ClassData data)
        {
            if (data.Students.Count == 0) return 80m;
            Dictionary<string, int> counts = new Dictionary<string, int>();
            Dictionary<string, decimal> values = new Dictionary<string, decimal>();
            for (int si = 0; si < data.Students.Count; si++)
            {
                Student s = data.Students[si];
                for (int m = 0; m < data.Modules.Count; m++)
                {
                    ModuleData md = s.Modules[m];
                    if (!md.HasScore) continue;
                    decimal residue = Math.Round(md.Score - md.SumDelta(), 2);
                    string key = residue.ToString(CultureInfo.InvariantCulture);
                    int c;
                    counts.TryGetValue(key, out c);
                    counts[key] = c + 1;
                    values[key] = residue;
                }
            }
            string bestKey = null;
            int bestCount = -1;
            foreach (KeyValuePair<string, int> kv in counts)
            {
                if (kv.Value > bestCount)
                {
                    bestCount = kv.Value;
                    bestKey = kv.Key;
                }
            }
            if (bestKey == null) return 80m;
            decimal best = values[bestKey];
            if (best < 0m || best > 1000m) return 80m;
            return best;
        }

        private static string InferClassName(string path)
        {
            string fn = Path.GetFileNameWithoutExtension(path);
            Match m = Regex.Match(fn, @"([A-Za-z0-9\u4e00-\u9fa5]{1,24}班)");
            if (m.Success) return m.Groups[1].Value;
            return fn;
        }

        private static Dictionary<string, string> ReadCustomProps(ZipArchive zip)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            string xml = XlsxCommon.ReadEntry(zip, "docProps/custom.xml");
            if (string.IsNullOrEmpty(xml)) return dict;
            try
            {
                XmlDocument doc = XlsxCommon.LoadXml(xml);
                XmlNodeList props = doc.SelectNodes("//*[local-name()='property']");
                if (props != null)
                {
                    foreach (XmlNode n in props)
                    {
                        XmlElement e = n as XmlElement;
                        if (e == null) continue;
                        string name = e.GetAttribute("name");
                        if (string.IsNullOrEmpty(name)) continue;
                        dict[name] = e.InnerText;
                    }
                }
            }
            catch { }
            return dict;
        }

        // ---------- 通用工作簿网格读取（用于导入名单） ----------

        public static List<string[]> ReadGrid(string path)
        {
            if (!File.Exists(path)) throw new XlsxException("文件不存在：" + path);
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read, false, Encoding.UTF8))
                {
                    string sheetPath = XlsxCommon.FindSheetPath(zip);
                    string sheetXml = XlsxCommon.ReadEntry(zip, sheetPath);
                    if (string.IsNullOrEmpty(sheetXml)) return new List<string[]>();
                    string dummy;
                    List<string> shared = XlsxCommon.ReadShared(zip, out dummy);
                    XmlDocument sheet = XlsxCommon.LoadXml(sheetXml);
                    XmlNamespaceManager ns = XlsxCommon.Ns(sheet);
                    List<RowInfo> rows = ReadRows(sheet, ns, shared);
                    int maxCol = 1;
                    foreach (RowInfo r in rows) if (r.MaxColumn > maxCol) maxCol = r.MaxColumn;
                    if (maxCol > 200) maxCol = 200;
                    List<string[]> result = new List<string[]>();
                    foreach (RowInfo r in rows)
                    {
                        string[] arr = new string[maxCol];
                        for (int c = 1; c <= maxCol; c++) arr[c - 1] = r.Get(c).Trim();
                        bool any = false;
                        for (int i = 0; i < arr.Length; i++) if (arr[i].Length > 0) { any = true; break; }
                        if (any) result.Add(arr);
                    }
                    return result;
                }
            }
            catch (XlsxException) { throw; }
            catch (Exception ex)
            {
                throw new XlsxException("无法读取名单文件：" + ex.Message, ex);
            }
        }

        // ---------- 内部行/单元格模型 ----------

        private class RowInfo
        {
            public int Row;
            public Dictionary<int, XmlElement> Cells = new Dictionary<int, XmlElement>();
            public int MaxColumn;

            public XmlElement GetElement(int col)
            {
                XmlElement e;
                return Cells.TryGetValue(col, out e) ? e : null;
            }

            public string Get(int col)
            {
                XmlElement e;
                if (!Cells.TryGetValue(col, out e)) return "";
                return _sharedText(e);
            }

            public Func<XmlElement, string> _textProvider;
            private string _sharedText(XmlElement e)
            {
                return _textProvider == null ? "" : _textProvider(e);
            }

            public bool HasText(string text)
            {
                foreach (KeyValuePair<int, XmlElement> kv in Cells)
                {
                    if (Get(kv.Key).Trim() == text) return true;
                }
                return false;
            }

            public int FindColumn(string text)
            {
                int best = 0;
                foreach (KeyValuePair<int, XmlElement> kv in Cells)
                {
                    if (Get(kv.Key).Trim() == text && (best == 0 || kv.Key < best)) best = kv.Key;
                }
                return best;
            }
        }

        private static List<RowInfo> ReadRows(XmlDocument sheet, XmlNamespaceManager ns, List<string> shared)
        {
            List<RowInfo> list = new List<RowInfo>();
            XmlNodeList rowNodes = sheet.SelectNodes("/m:worksheet/m:sheetData/m:row", ns);
            if (rowNodes == null) return list;
            foreach (XmlNode rn in rowNodes)
            {
                XmlElement rowEl = rn as XmlElement;
                if (rowEl == null) continue;
                RowInfo info = new RowInfo();
                int rowIndex;
                if (!int.TryParse(rowEl.GetAttribute("r"), out rowIndex)) continue;
                info.Row = rowIndex;
                info._textProvider = delegate(XmlElement e) { return XlsxCommon.CellText(e, shared); };
                foreach (XmlNode cn in rowEl.ChildNodes)
                {
                    XmlElement c = cn as XmlElement;
                    if (c == null || c.LocalName != "c") continue;
                    string colName = XlsxCommon.CellColumn(c.GetAttribute("r"));
                    int col = XlsxCommon.ColIndex(colName);
                    if (col <= 0) continue;
                    info.Cells[col] = c;
                    if (col > info.MaxColumn) info.MaxColumn = col;
                }
                list.Add(info);
            }
            list.Sort(delegate(RowInfo a, RowInfo b) { return a.Row.CompareTo(b.Row); });
            return list;
        }

        private static List<string> ReadMergeRefs(XmlDocument sheet, XmlNamespaceManager ns)
        {
            List<string> list = new List<string>();
            XmlNodeList nodes = sheet.SelectNodes("/m:worksheet/m:mergeCells/m:mergeCell", ns);
            if (nodes != null)
            {
                foreach (XmlNode n in nodes)
                {
                    XmlElement e = n as XmlElement;
                    if (e != null) list.Add(e.GetAttribute("ref"));
                }
            }
            return list;
        }

        private static bool TryParseRange(string range, out int c1, out int r1, out int c2, out int r2)
        {
            c1 = r1 = c2 = r2 = 0;
            if (string.IsNullOrEmpty(range)) return false;
            string[] parts = range.Split(':');
            if (parts.Length != 2) return false;
            c1 = XlsxCommon.ColIndex(XlsxCommon.CellColumn(parts[0]));
            r1 = XlsxCommon.CellRow(parts[0]);
            c2 = XlsxCommon.ColIndex(XlsxCommon.CellColumn(parts[1]));
            r2 = XlsxCommon.CellRow(parts[1]);
            return c1 > 0 && r1 > 0 && c2 > 0 && r2 > 0;
        }
    }
}
