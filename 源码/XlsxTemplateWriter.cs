using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace CaoXing
{
    /// <summary>在原模板副本上直接修改单元格，最大限度保留原有格式、合并、列宽等。</summary>
    public static class XlsxTemplateWriter
    {
        /// <summary>把数据写入 sourcePath 的副本 targetPath（targetPath 会被覆盖）。</summary>
        public static void Write(ClassData data, string sourcePath, string targetPath, bool rosterDirty)
        {
            if (!File.Exists(sourcePath)) throw new XlsxException("原表格不存在，无法保存：" + sourcePath);
            File.Copy(sourcePath, targetPath, true);
            try
            {
                using (TemplateEditor editor = new TemplateEditor(targetPath))
                {
                    editor.Update(data, rosterDirty);
                    editor.Commit();
                }
            }
            catch
            {
                try { File.Delete(targetPath); } catch { }
                throw;
            }
        }

        private class TemplateEditor : IDisposable
        {
            private FileStream _fs;
            private ZipArchive _zip;
            private XmlDocument _sheet;
            private XmlDocument _workbook;
            private XmlDocument _rootRels;
            private XmlDocument _contentTypes;
            private XmlDocument _customDoc;
            private string _sheetPath;
            private XmlNamespaceManager _sheetNs;

            public TemplateEditor(string path)
            {
                _fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                _zip = new ZipArchive(_fs, ZipArchiveMode.Update);
                _sheetPath = XlsxCommon.FindSheetPath(_zip);
                string sheetXml = XlsxCommon.ReadEntry(_zip, _sheetPath);
                if (string.IsNullOrEmpty(sheetXml)) throw new XlsxException("模板工作表为空。");
                _sheet = XlsxCommon.LoadXml(sheetXml);
                _sheetNs = XlsxCommon.Ns(_sheet);


                string wbXml = XlsxCommon.ReadEntry(_zip, "xl/workbook.xml");
                _workbook = string.IsNullOrEmpty(wbXml) ? null : XlsxCommon.LoadXml(wbXml);
                string relXml = XlsxCommon.ReadEntry(_zip, "_rels/.rels");
                _rootRels = string.IsNullOrEmpty(relXml) ? null : XlsxCommon.LoadXml(relXml);
                string ctXml = XlsxCommon.ReadEntry(_zip, "[Content_Types].xml");
                _contentTypes = string.IsNullOrEmpty(ctXml) ? null : XlsxCommon.LoadXml(ctXml);
                string customXml = XlsxCommon.ReadEntry(_zip, "docProps/custom.xml");
                _customDoc = string.IsNullOrEmpty(customXml) ? null : XlsxCommon.LoadXml(customXml);
            }

            public void Update(ClassData data, bool rosterDirty)
            {
                if (data.ScoreColumns == null || data.ScoreColumns.Length != data.ModuleCount)
                    throw new XlsxException("模块列信息不完整，无法保存。");

                Dictionary<int, string> styles = CaptureStyles(data.DataStartRow);

                if (rosterDirty)
                {
                    RebuildRows(data, styles);
                    RemoveElement(_sheet.DocumentElement, "sortState");
                }
                else
                {
                    for (int i = 0; i < data.Students.Count; i++)
                    {
                        Student s = data.Students[i];
                        XmlElement row = FindRow(s.RowIndex);
                        if (row == null) continue;
                        bool changed = false;
                        for (int m = 0; m < data.ModuleCount; m++)
                        {
                            ModuleData md = s.Modules[m];
                            if (!md.Dirty) continue;
                            string scoreStyle = GetStyle(styles, data.ScoreColumns[m]);
                            string detailStyle = GetStyle(styles, data.DetailColumns[m]);
                            SetNumber(GetOrCreateCell(row, data.ScoreColumns[m], scoreStyle), data.GetModuleScore(s, m));
                            SetText(GetOrCreateCell(row, data.DetailColumns[m], detailStyle), data.DetailText(s.Modules[m]));
                            changed = true;
                        }
                        if (changed) SetAverageCell(row, data, s, GetStyle(styles, data.AverageColumn));
                    }
                }

                UpdateCustomProperties(data);

                if (_workbook != null)
                {
                    XmlElement calcPr = XlsxCommon.Child(_workbook.DocumentElement, "calcPr");
                    if (calcPr == null)
                    {
                        calcPr = _workbook.CreateElement("calcPr", XlsxCommon.NsMain);
                        XmlElement ext = XlsxCommon.Child(_workbook.DocumentElement, "extLst");
                        if (ext != null) _workbook.DocumentElement.InsertBefore(calcPr, ext);
                        else _workbook.DocumentElement.AppendChild(calcPr);
                    }
                    calcPr.SetAttribute("calcId", "191029");
                    calcPr.SetAttribute("fullCalcOnLoad", "1");
                }

                UpdateDimension(data);
            }

            private Dictionary<int, string> CaptureStyles(int dataStart)
            {
                Dictionary<int, string> styles = new Dictionary<int, string>();
                XmlNodeList rows = _sheet.SelectNodes("/m:worksheet/m:sheetData/m:row", _sheetNs);
                if (rows == null) return styles;
                foreach (XmlNode rn in rows)
                {
                    XmlElement row = rn as XmlElement;
                    if (row == null) continue;
                    int rowIndex;
                    if (!int.TryParse(row.GetAttribute("r"), out rowIndex) || rowIndex < dataStart) continue;
                    foreach (XmlNode cn in row.ChildNodes)
                    {
                        XmlElement c = cn as XmlElement;
                        if (c == null || c.LocalName != "c") continue;
                        int col = XlsxCommon.ColIndex(XlsxCommon.CellColumn(c.GetAttribute("r")));
                        if (col <= 0) continue;
                        string s = c.GetAttribute("s");
                        if (!string.IsNullOrEmpty(s) && !styles.ContainsKey(col)) styles[col] = s;
                    }
                }
                return styles;
            }

            private static string GetStyle(Dictionary<int, string> styles, int col)
            {
                string s;
                return styles.TryGetValue(col, out s) ? s : "";
            }

            private void RebuildRows(ClassData data, Dictionary<int, string> styles)
            {
                int start = data.DataStartRow;
                int oldLast = start - 1;
                for (int i = 0; i < data.Students.Count; i++)
                    if (data.Students[i].RowIndex > oldLast) oldLast = data.Students[i].RowIndex;
                int newLast = start + data.Students.Count - 1;
                int removeEnd = Math.Max(oldLast, newLast);
                RemoveRows(start, removeEnd);

                for (int i = 0; i < data.Students.Count; i++)
                {
                    Student s = data.Students[i];
                    int rowIndex = start + i;
                    s.RowIndex = rowIndex;
                    if (s.Seq <= 0) s.Seq = i + 1;
                    XmlElement row = CreateRow(rowIndex, data.AverageColumn);
                    SetNumber(GetOrCreateCell(row, 1, GetStyle(styles, 1)), s.Seq);
                    SetText(GetOrCreateCell(row, 2, GetStyle(styles, 2)), s.No);
                    SetText(GetOrCreateCell(row, 3, GetStyle(styles, 3)), s.Name);
                    for (int m = 0; m < data.ModuleCount; m++)
                    {
                        SetNumber(GetOrCreateCell(row, data.ScoreColumns[m], GetStyle(styles, data.ScoreColumns[m])), data.GetModuleScore(s, m));
                        SetText(GetOrCreateCell(row, data.DetailColumns[m], GetStyle(styles, data.DetailColumns[m])), data.DetailText(s.Modules[m]));
                    }
                    SetAverageCell(row, data, s, GetStyle(styles, data.AverageColumn));
                    InsertRow(row);
                }
            }

            private void SetAverageCell(XmlElement row, ClassData data, Student s, string style)
            {
                XmlElement cell = GetOrCreateCell(row, data.AverageColumn, style);
                StringBuilder formula = new StringBuilder("AVERAGE(");
                for (int m = 0; m < data.ModuleCount; m++)
                {
                    if (m > 0) formula.Append(',');
                    formula.Append(XlsxCommon.CellRef(data.ScoreColumns[m], s.RowIndex));
                }
                formula.Append(')');
                SetFormula(cell, formula.ToString(), data.GetAverage(s));
            }

            private void UpdateDimension(ClassData data)
            {
                int lastRow = 2;
                XmlNodeList rows = _sheet.SelectNodes("/m:worksheet/m:sheetData/m:row", _sheetNs);
                if (rows != null)
                {
                    foreach (XmlNode n in rows)
                    {
                        XmlElement e = n as XmlElement;
                        int r;
                        if (e != null && int.TryParse(e.GetAttribute("r"), out r) && r > lastRow) lastRow = r;
                    }
                }
                for (int i = 0; i < data.Students.Count; i++)
                    if (data.Students[i].RowIndex > lastRow) lastRow = data.Students[i].RowIndex;
                XmlElement dim = _sheet.SelectSingleNode("/m:worksheet/m:dimension", _sheetNs) as XmlElement;
                if (dim == null)
                {
                    dim = _sheet.CreateElement("dimension", XlsxCommon.NsMain);
                    _sheet.DocumentElement.PrependChild(dim);
                }
                dim.SetAttribute("ref", "A1:" + XlsxCommon.CellRef(data.AverageColumn, lastRow));
            }

            private void UpdateCustomProperties(ClassData data)
            {
                EnsureCustomPart();
                SetCustom("CX_AppVersion", ClassData.AppVersion);
                SetCustom("CX_ClassName", data.ClassName);
                SetCustom("CX_ModuleCount", data.ModuleCount.ToString(CultureInfo.InvariantCulture));
                StringBuilder aspects = new StringBuilder();
                StringBuilder names = new StringBuilder();
                for (int i = 0; i < data.ModuleCount; i++)
                {
                    if (i > 0) { aspects.Append(','); names.Append('|'); }
                    aspects.Append(data.Modules[i].AspectCount.ToString(CultureInfo.InvariantCulture));
                    names.Append(data.Modules[i].Name.Replace('|', '/'));
                }
                SetCustom("CX_AspectCounts", aspects.ToString());
                SetCustom("CX_ModuleNames", names.ToString());
                SetCustom("CX_InitialScore", Format.Decimal(data.InitialScore));
                SetCustom("CX_LastSaved", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            }

            private void EnsureCustomPart()
            {
                if (_customDoc != null) return;
                _customDoc = new XmlDocument();
                _customDoc.PreserveWhitespace = true;
                XmlElement root = _customDoc.CreateElement("Properties", XlsxCommon.NsCustom);
                XmlAttribute vt = _customDoc.CreateAttribute("xmlns", "vt", "http://www.w3.org/2000/xmlns/");
                vt.Value = XlsxCommon.NsVt;
                root.Attributes.Append(vt);
                _customDoc.AppendChild(root);

                if (_contentTypes != null)
                {
                    XmlNamespaceManager ctNs = XlsxCommon.Ns(_contentTypes);
                    bool exists = false;
                    foreach (XmlElement e in _contentTypes.SelectNodes("/ct:Types/ct:Override", ctNs))
                        if (e.GetAttribute("PartName") == "/docProps/custom.xml") { exists = true; break; }
                    if (!exists)
                    {
                        XmlElement ov = _contentTypes.CreateElement("Override", XlsxCommon.NsContentTypes);
                        ov.SetAttribute("PartName", "/docProps/custom.xml");
                        ov.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml");
                        _contentTypes.DocumentElement.AppendChild(ov);
                    }
                }

                if (_rootRels != null)
                {
                    XmlNamespaceManager relNs = XlsxCommon.Ns(_rootRels);
                    bool exists = false;
                    int maxId = 0;
                    foreach (XmlElement e in _rootRels.SelectNodes("/pr:Relationships/pr:Relationship", relNs))
                    {
                        if (e.GetAttribute("Type").EndsWith("/custom-properties")) exists = true;
                        string id = e.GetAttribute("Id");
                        if (id.StartsWith("rId"))
                        {
                            int n;
                            if (int.TryParse(id.Substring(3), out n) && n > maxId) maxId = n;
                        }
                    }
                    if (!exists)
                    {
                        XmlElement rel = _rootRels.CreateElement("Relationship", XlsxCommon.NsPkgRel);
                        rel.SetAttribute("Id", "rId" + (maxId + 1).ToString(CultureInfo.InvariantCulture));
                        rel.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
                        rel.SetAttribute("Target", "docProps/custom.xml");
                        _rootRels.DocumentElement.AppendChild(rel);
                    }
                }
            }

            private void SetCustom(string name, string value)
            {
                XmlElement root = _customDoc.DocumentElement;
                XmlElement found = null;
                int maxPid = 1;
                foreach (XmlNode n in root.ChildNodes)
                {
                    XmlElement e = n as XmlElement;
                    if (e == null || e.LocalName != "property") continue;
                    int pid;
                    if (int.TryParse(e.GetAttribute("pid"), out pid) && pid > maxPid) maxPid = pid;
                    if (e.GetAttribute("name") == name) found = e;
                }
                if (found == null)
                {
                    found = _customDoc.CreateElement("property", XlsxCommon.NsCustom);
                    found.SetAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
                    found.SetAttribute("pid", (maxPid + 1).ToString(CultureInfo.InvariantCulture));
                    found.SetAttribute("name", name);
                    root.AppendChild(found);
                }
                while (found.HasChildNodes) found.RemoveChild(found.FirstChild);
                XmlElement v = _customDoc.CreateElement("vt", "lpwstr", XlsxCommon.NsVt);
                v.InnerText = value ?? "";
                found.AppendChild(v);
            }

            // ---------- 单元格与行操作 ----------

            private XmlElement FindRow(int rowIndex)
            {
                return _sheet.SelectSingleNode(string.Format("/m:worksheet/m:sheetData/m:row[@r='{0}']", rowIndex), _sheetNs) as XmlElement;
            }

            private XmlElement CreateRow(int rowIndex, int lastCol)
            {
                XmlElement row = _sheet.CreateElement("row", XlsxCommon.NsMain);
                row.SetAttribute("r", rowIndex.ToString(CultureInfo.InvariantCulture));
                row.SetAttribute("spans", "1:" + lastCol.ToString(CultureInfo.InvariantCulture));
                return row;
            }

            private void InsertRow(XmlElement row)
            {
                XmlElement sheetData = _sheet.SelectSingleNode("/m:worksheet/m:sheetData", _sheetNs) as XmlElement;
                if (sheetData == null) return;
                int rowIndex;
                int.TryParse(row.GetAttribute("r"), out rowIndex);
                XmlNode insertBefore = null;
                foreach (XmlNode n in sheetData.ChildNodes)
                {
                    XmlElement e = n as XmlElement;
                    if (e == null || e.LocalName != "row") continue;
                    int r;
                    if (int.TryParse(e.GetAttribute("r"), out r) && r > rowIndex) { insertBefore = e; break; }
                }
                if (insertBefore != null) sheetData.InsertBefore(row, insertBefore);
                else sheetData.AppendChild(row);
            }

            private void RemoveRows(int start, int end)
            {
                if (end < start) return;
                XmlElement sheetData = _sheet.SelectSingleNode("/m:worksheet/m:sheetData", _sheetNs) as XmlElement;
                if (sheetData == null) return;
                List<XmlElement> remove = new List<XmlElement>();
                foreach (XmlNode n in sheetData.ChildNodes)
                {
                    XmlElement e = n as XmlElement;
                    if (e == null || e.LocalName != "row") continue;
                    int r;
                    if (int.TryParse(e.GetAttribute("r"), out r) && r >= start && r <= end) remove.Add(e);
                }
                for (int i = 0; i < remove.Count; i++) sheetData.RemoveChild(remove[i]);
            }

            private XmlElement GetOrCreateCell(XmlElement row, int col, string style)
            {
                int rowIndex;
                int.TryParse(row.GetAttribute("r"), out rowIndex);
                string cellRef = XlsxCommon.CellRef(col, rowIndex);
                foreach (XmlNode n in row.ChildNodes)
                {
                    XmlElement e = n as XmlElement;
                    if (e != null && e.LocalName == "c" && e.GetAttribute("r") == cellRef) return e;
                }
                XmlElement cell = _sheet.CreateElement("c", XlsxCommon.NsMain);
                cell.SetAttribute("r", cellRef);
                if (!string.IsNullOrEmpty(style)) cell.SetAttribute("s", style);
                XmlNode insertBefore = null;
                foreach (XmlNode n in row.ChildNodes)
                {
                    XmlElement e = n as XmlElement;
                    if (e == null || e.LocalName != "c") continue;
                    int c = XlsxCommon.ColIndex(XlsxCommon.CellColumn(e.GetAttribute("r")));
                    if (c > col) { insertBefore = e; break; }
                }
                if (insertBefore != null) row.InsertBefore(cell, insertBefore);
                else row.AppendChild(cell);
                return cell;
            }

            private void SetText(XmlElement cell, string text)
            {
                if (text == null) text = "";
                RemoveChild(cell, "f");
                RemoveChild(cell, "v");
                RemoveChild(cell, "is");
                if (text.Length == 0)
                {
                    if (cell.HasAttribute("t")) cell.RemoveAttribute("t");
                    return;
                }
                cell.SetAttribute("t", "inlineStr");
                XmlElement isEl = _sheet.CreateElement("is", XlsxCommon.NsMain);
                XmlElement t = _sheet.CreateElement("t", XlsxCommon.NsMain);
                if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
                    t.SetAttribute("space", XlsxCommon.NsXml, "preserve");
                t.InnerText = text;
                isEl.AppendChild(t);
                cell.AppendChild(isEl);
            }            private void SetNumber(XmlElement cell, decimal value)
            {
                if (cell.HasAttribute("t")) cell.RemoveAttribute("t");
                RemoveChild(cell, "f");
                RemoveChild(cell, "is");
                XmlElement v = XlsxCommon.Child(cell, "v");
                if (v == null)
                {
                    v = _sheet.CreateElement("v", XlsxCommon.NsMain);
                    cell.AppendChild(v);
                }
                v.InnerText = Format.Raw(value);
            }

            private void SetFormula(XmlElement cell, string formula, decimal cached)
            {
                if (cell.HasAttribute("t")) cell.RemoveAttribute("t");
                RemoveChild(cell, "is");
                XmlElement f = XlsxCommon.Child(cell, "f");
                if (f == null)
                {
                    f = _sheet.CreateElement("f", XlsxCommon.NsMain);
                    cell.InsertBefore(f, cell.FirstChild);
                }
                f.InnerText = formula;
                XmlElement v = XlsxCommon.Child(cell, "v");
                if (v == null)
                {
                    v = _sheet.CreateElement("v", XlsxCommon.NsMain);
                    cell.AppendChild(v);
                }
                v.InnerText = Format.Raw(cached);
            }

            private static void RemoveChild(XmlElement parent, string localName)
            {
                XmlElement e = XlsxCommon.Child(parent, localName);
                if (e != null) parent.RemoveChild(e);
            }

            private static void RemoveElement(XmlElement parent, string localName)
            {
                if (parent == null) return;
                XmlElement e = XlsxCommon.Child(parent, localName);
                if (e != null) parent.RemoveChild(e);
            }

            public void Commit()
            {
                WriteEntry(_sheetPath, XlsxCommon.XmlBytes(_sheet));
                if (_workbook != null) WriteEntry("xl/workbook.xml", XlsxCommon.XmlBytes(_workbook));
                if (_contentTypes != null) WriteEntry("[Content_Types].xml", XlsxCommon.XmlBytes(_contentTypes));
                if (_rootRels != null) WriteEntry("_rels/.rels", XlsxCommon.XmlBytes(_rootRels));
                if (_customDoc != null) WriteEntry("docProps/custom.xml", XlsxCommon.XmlBytes(_customDoc));
            }

            private void WriteEntry(string name, byte[] bytes)
            {
                ZipArchiveEntry e = _zip.GetEntry(name);
                if (e != null) e.Delete();
                ZipArchiveEntry ne = _zip.CreateEntry(name, CompressionLevel.Optimal);
                using (Stream s = ne.Open())
                {
                    s.Write(bytes, 0, bytes.Length);
                }
            }

            public void Dispose()
            {
                if (_zip != null) { _zip.Dispose(); _zip = null; }
                if (_fs != null) { _fs.Dispose(); _fs = null; }
            }
        }
    }
}


