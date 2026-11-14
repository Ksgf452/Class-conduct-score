using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace CaoXing
{
    /// <summary>从零生成操行汇总表（用于新建班级或非模板数据）。</summary>
    public static class XlsxCreator
    {
        private const string NsMain = XlsxCommon.NsMain;
        private const string NsRel = XlsxCommon.NsRel;

        public static void Create(ClassData data, string targetPath)
        {
            int moduleCount = data.ModuleCount;
            if (moduleCount <= 0) throw new XlsxException("模块数量必须大于 0。");
            int avgCol = 4 + moduleCount * 2;
            int lastCol = avgCol;
            int lastRow = 2 + data.Students.Count;

            List<string> strings = new List<string>();
            Dictionary<string, int> map = new Dictionary<string, int>(StringComparer.Ordinal);
            AddString(strings, map, "序号");
            AddString(strings, map, "学号");
            AddString(strings, map, "姓名");
            for (int m = 0; m < moduleCount; m++) AddString(strings, map, data.Modules[m].Name);
            AddString(strings, map, "平均分");
            for (int i = 0; i < data.Students.Count; i++)
            {
                AddString(strings, map, data.Students[i].No);
                AddString(strings, map, data.Students[i].Name);
                for (int m = 0; m < moduleCount; m++)
                    AddString(strings, map, data.DetailText(data.Students[i].Modules[m]));
            }

            string sheetXml = BuildSheet(data, strings, map, avgCol, lastCol, lastRow);
            string sharedXml = BuildSharedStrings(strings);
            string stylesXml = StylesXml;
            string workbookXml = BuildWorkbook();
            string customXml = BuildCustomProps(data);
            string coreXml = BuildCore(data);
            string appXml = BuildApp();
            string contentTypes = BuildContentTypes();
            string rootRels = BuildRootRels();
            string workbookRels = BuildWorkbookRels();

            using (FileStream fs = new FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                AddText(zip, "[Content_Types].xml", contentTypes);
                AddText(zip, "_rels/.rels", rootRels);
                AddText(zip, "docProps/core.xml", coreXml);
                AddText(zip, "docProps/app.xml", appXml);
                AddText(zip, "docProps/custom.xml", customXml);
                AddText(zip, "xl/workbook.xml", workbookXml);
                AddText(zip, "xl/_rels/workbook.xml.rels", workbookRels);
                AddText(zip, "xl/styles.xml", stylesXml);
                AddText(zip, "xl/sharedStrings.xml", sharedXml);
                AddText(zip, "xl/worksheets/sheet1.xml", sheetXml);
            }
        }

        private static void AddString(List<string> list, Dictionary<string, int> map, string s)
        {
            if (s == null) s = "";
            if (!map.ContainsKey(s))
            {
                map[s] = list.Count;
                list.Add(s);
            }
        }

        private static void AddText(ZipArchive zip, string name, string text)
        {
            ZipArchiveEntry e = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (Stream s = e.Open())
            {
                byte[] bytes = new UTF8Encoding(false).GetBytes(text);
                s.Write(bytes, 0, bytes.Length);
            }
        }

        private static string BuildSheet(ClassData data, List<string> strings, Dictionary<string, int> map,
            int avgCol, int lastCol, int lastRow)
        {
            MemoryStream ms = new MemoryStream();
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UTF8Encoding(false);
            settings.Indent = false;
            settings.OmitXmlDeclaration = false;

            using (XmlWriter w = XmlWriter.Create(ms, settings))
            {
                w.WriteStartDocument(true);
                w.WriteStartElement("worksheet", NsMain);
                w.WriteAttributeString("xmlns", "r", null, NsRel);
                w.WriteStartElement("dimension", NsMain);
                w.WriteAttributeString("ref", "A1:" + XlsxCommon.CellRef(avgCol, lastRow));
                w.WriteEndElement();

                w.WriteStartElement("sheetViews", NsMain);
                w.WriteStartElement("sheetView", NsMain);
                w.WriteAttributeString("tabSelected", "1");
                w.WriteAttributeString("zoomScale", "85");
                w.WriteAttributeString("zoomScaleNormal", "85");
                w.WriteAttributeString("workbookViewId", "0");
                w.WriteStartElement("pane", NsMain);
                w.WriteAttributeString("xSplit", "3");
                w.WriteAttributeString("ySplit", "2");
                w.WriteAttributeString("topLeftCell", "D3");
                w.WriteAttributeString("activePane", "bottomRight");
                w.WriteAttributeString("state", "frozen");
                w.WriteEndElement();
                w.WriteStartElement("selection", NsMain);
                w.WriteAttributeString("pane", "bottomRight");
                w.WriteAttributeString("activeCell", "D3");
                w.WriteAttributeString("sqref", "D3");
                w.WriteEndElement();
                w.WriteEndElement();
                w.WriteEndElement();

                w.WriteStartElement("sheetFormatPr", NsMain);
                w.WriteAttributeString("defaultColWidth", "9");
                w.WriteAttributeString("defaultRowHeight", "18");
                w.WriteEndElement();

                w.WriteStartElement("cols", NsMain);
                WriteCol(w, 1, 1, 5.5);
                WriteCol(w, 2, 2, 14);
                WriteCol(w, 3, 3, 10);
                for (int m = 0; m < data.ModuleCount; m++)
                {
                    WriteCol(w, 4 + m * 2, 4 + m * 2, 7.5);
                    WriteCol(w, 5 + m * 2, 5 + m * 2, 48);
                }
                WriteCol(w, avgCol, avgCol, 10);
                w.WriteEndElement();

                w.WriteStartElement("sheetData", NsMain);
                // 表头第一行
                WriteRowStart(w, 1, lastCol, 30);
                WriteCellText(w, 1, 1, 1, "序号", map);
                WriteCellText(w, 2, 1, 1, "学号", map);
                WriteCellText(w, 3, 1, 1, "姓名", map);
                for (int m = 0; m < data.ModuleCount; m++)
                {
                    WriteCellText(w, 4 + m * 2, 1, 1, data.Modules[m].Name, map);
                    WriteCellEmpty(w, 5 + m * 2, 1, 1);
                }
                WriteCellText(w, avgCol, 1, 1, "平均分", map);
                WriteRowEnd(w);

                // 表头第二行（用于合并）
                WriteRowStart(w, 2, lastCol, 22);
                for (int c = 1; c <= lastCol; c++) WriteCellEmpty(w, c, 2, 1);
                WriteRowEnd(w);

                for (int i = 0; i < data.Students.Count; i++)
                {
                    Student s = data.Students[i];
                    int row = 3 + i;
                    int height = EstimateRowHeight(data, s);
                    WriteRowStart(w, row, lastCol, height);
                    WriteCellNumber(w, 1, row, 2, s.Seq);
                    WriteCellText(w, 2, row, 2, s.No, map);
                    WriteCellText(w, 3, row, 2, s.Name, map);
                    for (int m = 0; m < data.ModuleCount; m++)
                    {
                        WriteCellNumber(w, 4 + m * 2, row, 3, data.GetModuleScore(s, m));
                        WriteCellText(w, 5 + m * 2, row, 4, data.DetailText(s.Modules[m]), map);
                    }
                    StringBuilder formula = new StringBuilder("AVERAGE(");
                    for (int m = 0; m < data.ModuleCount; m++)
                    {
                        if (m > 0) formula.Append(',');
                        formula.Append(XlsxCommon.CellRef(4 + m * 2, row));
                    }
                    formula.Append(')');
                    WriteCellFormula(w, avgCol, row, 5, formula.ToString(), data.GetAverage(s));
                    WriteRowEnd(w);
                }
                w.WriteEndElement(); // sheetData

                w.WriteStartElement("mergeCells", NsMain);
                w.WriteAttributeString("count", (4 + data.ModuleCount).ToString(CultureInfo.InvariantCulture));
                WriteMerge(w, "A1:A2");
                WriteMerge(w, "B1:B2");
                WriteMerge(w, "C1:C2");
                for (int m = 0; m < data.ModuleCount; m++)
                {
                    string c1 = XlsxCommon.ColName(4 + m * 2);
                    string c2 = XlsxCommon.ColName(5 + m * 2);
                    WriteMerge(w, c1 + "1:" + c2 + "2");
                }
                WriteMerge(w, XlsxCommon.ColName(avgCol) + "1:" + XlsxCommon.ColName(avgCol) + "2");
                w.WriteEndElement();

                w.WriteStartElement("pageMargins", NsMain);
                w.WriteAttributeString("left", "0.7");
                w.WriteAttributeString("right", "0.7");
                w.WriteAttributeString("top", "0.75");
                w.WriteAttributeString("bottom", "0.75");
                w.WriteAttributeString("header", "0.3");
                w.WriteAttributeString("footer", "0.3");
                w.WriteEndElement();

                w.WriteEndElement(); // worksheet
                w.WriteEndDocument();
            }
            return Encoding.UTF8.GetString(ms.ToArray());
        }

        private static int EstimateRowHeight(ClassData data, Student s)
        {
            int lines = 1;
            for (int m = 0; m < data.ModuleCount; m++)
            {
                string text = data.DetailText(s.Modules[m]);
                int l = 1;
                string[] parts = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
                for (int i = 0; i < parts.Length; i++)
                {
                    int n = (int)Math.Ceiling(parts[i].Length / 24.0);
                    if (n < 1) n = 1;
                    l += n;
                }
                if (l > lines) lines = l;
            }
            int h = 10 + lines * 15;
            if (h < 28) h = 28;
            if (h > 180) h = 180;
            return h;
        }

        private static void WriteCol(XmlWriter w, int min, int max, double width)
        {
            w.WriteStartElement("col", NsMain);
            w.WriteAttributeString("min", min.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("max", max.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("width", width.ToString("0.##", CultureInfo.InvariantCulture));
            w.WriteAttributeString("customWidth", "1");
            w.WriteEndElement();
        }

        private static void WriteRowStart(XmlWriter w, int row, int lastCol, int height)
        {
            w.WriteStartElement("row", NsMain);
            w.WriteAttributeString("r", row.ToString(CultureInfo.InvariantCulture));
            w.WriteAttributeString("spans", "1:" + lastCol.ToString(CultureInfo.InvariantCulture));
            if (height > 0)
            {
                w.WriteAttributeString("ht", height.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("customHeight", "1");
            }
        }

        private static void WriteRowEnd(XmlWriter w)
        {
            w.WriteEndElement();
        }

        private static void WriteCellText(XmlWriter w, int col, int row, int style, string text, Dictionary<string, int> map)
        {
            if (text == null) text = "";
            w.WriteStartElement("c", NsMain);
            w.WriteAttributeString("r", XlsxCommon.CellRef(col, row));
            w.WriteAttributeString("s", style.ToString(CultureInfo.InvariantCulture));
            if (text.Length > 0)
            {
                int idx;
                if (!map.TryGetValue(text, out idx)) idx = 0;
                w.WriteAttributeString("t", "s");
                w.WriteStartElement("v", NsMain);
                w.WriteString(idx.ToString(CultureInfo.InvariantCulture));
                w.WriteEndElement();
            }
            w.WriteEndElement();
        }

        private static void WriteCellEmpty(XmlWriter w, int col, int row, int style)
        {
            w.WriteStartElement("c", NsMain);
            w.WriteAttributeString("r", XlsxCommon.CellRef(col, row));
            w.WriteAttributeString("s", style.ToString(CultureInfo.InvariantCulture));
            w.WriteEndElement();
        }

        private static void WriteCellNumber(XmlWriter w, int col, int row, int style, decimal value)
        {
            w.WriteStartElement("c", NsMain);
            w.WriteAttributeString("r", XlsxCommon.CellRef(col, row));
            w.WriteAttributeString("s", style.ToString(CultureInfo.InvariantCulture));
            w.WriteStartElement("v", NsMain);
            w.WriteString(Format.Raw(value));
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteCellFormula(XmlWriter w, int col, int row, int style, string formula, decimal cached)
        {
            w.WriteStartElement("c", NsMain);
            w.WriteAttributeString("r", XlsxCommon.CellRef(col, row));
            w.WriteAttributeString("s", style.ToString(CultureInfo.InvariantCulture));
            w.WriteStartElement("f", NsMain);
            w.WriteString(formula);
            w.WriteEndElement();
            w.WriteStartElement("v", NsMain);
            w.WriteString(Format.Raw(cached));
            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static void WriteMerge(XmlWriter w, string range)
        {
            w.WriteStartElement("mergeCell", NsMain);
            w.WriteAttributeString("ref", range);
            w.WriteEndElement();
        }

        private static string BuildSharedStrings(List<string> strings)
        {
            MemoryStream ms = new MemoryStream();
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UTF8Encoding(false);
            settings.Indent = false;
            settings.OmitXmlDeclaration = false;
            using (XmlWriter w = XmlWriter.Create(ms, settings))
            {
                w.WriteStartDocument(true);
                w.WriteStartElement("sst", NsMain);
                w.WriteAttributeString("count", strings.Count.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("uniqueCount", strings.Count.ToString(CultureInfo.InvariantCulture));
                for (int i = 0; i < strings.Count; i++)
                {
                    w.WriteStartElement("si", NsMain);
                    w.WriteStartElement("t", NsMain);
                    if (strings[i].Length > 0 && (char.IsWhiteSpace(strings[i][0]) || char.IsWhiteSpace(strings[i][strings[i].Length - 1])))
                        w.WriteAttributeString("xml", "space", XlsxCommon.NsXml, "preserve");
                    w.WriteString(strings[i]);
                    w.WriteEndElement();
                    w.WriteEndElement();
                }
                w.WriteEndElement();
                w.WriteEndDocument();
            }
            return Encoding.UTF8.GetString(ms.ToArray());
        }

        private static string BuildWorkbook()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<workbook xmlns=\"").Append(NsMain).Append("\" xmlns:r=\"").Append(NsRel).Append("\">");
            sb.Append("<bookViews><workbookView windowWidth=\"24000\" windowHeight=\"12000\"/></bookViews>");
            sb.Append("<sheets><sheet name=\"操行汇总\" sheetId=\"1\" r:id=\"rId1\"/></sheets>");
            sb.Append("<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\"/>");
            sb.Append("</workbook>");
            return sb.ToString();
        }

        private static string BuildCustomProps(ClassData data)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Properties xmlns=\"").Append(XlsxCommon.NsCustom).Append("\" xmlns:vt=\"").Append(XlsxCommon.NsVt).Append("\">");
            Dictionary<string, string> props = new Dictionary<string, string>();
            props["CX_AppVersion"] = ClassData.AppVersion;
            props["CX_ClassName"] = data.ClassName;
            props["CX_ModuleCount"] = data.ModuleCount.ToString(CultureInfo.InvariantCulture);
            StringBuilder aspects = new StringBuilder();
            StringBuilder names = new StringBuilder();
            for (int i = 0; i < data.ModuleCount; i++)
            {
                if (i > 0) { aspects.Append(','); names.Append('|'); }
                aspects.Append(data.Modules[i].AspectCount.ToString(CultureInfo.InvariantCulture));
                names.Append(data.Modules[i].Name.Replace('|', '/'));
            }
            props["CX_AspectCounts"] = aspects.ToString();
            props["CX_ModuleNames"] = names.ToString();
            props["CX_InitialScore"] = Format.Decimal(data.InitialScore);
            props["CX_LastSaved"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            int pid = 2;
            foreach (KeyValuePair<string, string> kv in props)
            {
                sb.Append("<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"").Append(pid).Append("\" name=\"")
                  .Append(XlsxCommon.XmlEscape(kv.Key)).Append("\"><vt:lpwstr>").Append(XlsxCommon.XmlEscape(kv.Value))
                  .Append("</vt:lpwstr></property>");
                pid++;
            }
            sb.Append("</Properties>");
            return sb.ToString();
        }

        private static string BuildCore(ClassData data)
        {
            string now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" ");
            sb.Append("xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" ");
            sb.Append("xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            sb.Append("<dc:title>").Append(XlsxCommon.XmlEscape(data.ClassName)).Append(" 操行分汇总</dc:title>");
            sb.Append("<dc:creator>班级操行分管理系统</dc:creator>");
            sb.Append("<cp:lastModifiedBy>班级操行分管理系统</cp:lastModifiedBy>");
            sb.Append("<dcterms:created xsi:type=\"dcterms:W3CDTF\">").Append(now).Append("</dcterms:created>");
            sb.Append("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">").Append(now).Append("</dcterms:modified>");
            sb.Append("</cp:coreProperties>");
            return sb.ToString();
        }

        private static string BuildApp()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">" +
                   "<Application>班级操行分管理系统</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>" +
                   "<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>" +
                   "<HyperlinksChanged>false</HyperlinksChanged><AppVersion>1.0</AppVersion></Properties>";
        }

        private static string BuildContentTypes()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Types xmlns=\"").Append(XlsxCommon.NsContentTypes).Append("\">");
            sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            sb.Append("<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            sb.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
            sb.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            sb.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
            sb.Append("<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>");
            sb.Append("</Types>");
            return sb.ToString();
        }

        private static string BuildRootRels()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Relationships xmlns=\"").Append(XlsxCommon.NsPkgRel).Append("\">");
            sb.Append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>");
            sb.Append("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>");
            sb.Append("<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>");
            sb.Append("<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties\" Target=\"docProps/custom.xml\"/>");
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private static string BuildWorkbookRels()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Relationships xmlns=\"").Append(XlsxCommon.NsPkgRel).Append("\">");
            sb.Append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>");
            sb.Append("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            sb.Append("<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private const string StylesXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<styleSheet xmlns=\"" + NsMain + "\">" +
            "<fonts count=\"3\">" +
            "<font><sz val=\"11\"/><color rgb=\"FF1F2937\"/><name val=\"微软雅黑\"/><charset val=\"134\"/></font>" +
            "<font><b/><sz val=\"11\"/><color rgb=\"FFFFFFFF\"/><name val=\"微软雅黑\"/><charset val=\"134\"/></font>" +
            "<font><b/><sz val=\"11\"/><color rgb=\"FF1F4E79\"/><name val=\"微软雅黑\"/><charset val=\"134\"/></font>" +
            "</fonts>" +
            "<fills count=\"4\">" +
            "<fill><patternFill patternType=\"none\"/></fill>" +
            "<fill><patternFill patternType=\"gray125\"/></fill>" +
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF1F4E79\"/><bgColor indexed=\"64\"/></patternFill></fill>" +
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFEAF2FB\"/><bgColor indexed=\"64\"/></patternFill></fill>" +
            "</fills>" +
            "<borders count=\"2\">" +
            "<border><left/><right/><top/><bottom/><diagonal/></border>" +
            "<border><left style=\"thin\"><color rgb=\"FFB8C6D9\"/></left><right style=\"thin\"><color rgb=\"FFB8C6D9\"/></right>" +
            "<top style=\"thin\"><color rgb=\"FFB8C6D9\"/></top><bottom style=\"thin\"><color rgb=\"FFB8C6D9\"/></bottom><diagonal/></border>" +
            "</borders>" +
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
            "<cellXfs count=\"6\">" +
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"center\"/></xf>" +
            "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf>" +
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>" +
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>" +
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"left\" vertical=\"top\" wrapText=\"1\"/></xf>" +
            "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"3\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\"/></xf>" +
            "</cellXfs>" +
            "<cellStyles count=\"1\"><cellStyle name=\"常规\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
            "</styleSheet>";
    }
}

