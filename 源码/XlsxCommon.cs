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
    public class XlsxException : Exception
    {
        public XlsxException(string message) : base(message) { }
        public XlsxException(string message, Exception inner) : base(message, inner) { }
    }

    internal static class XlsxCommon
    {
        public const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public const string NsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        public const string NsContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        public const string NsCustom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        public const string NsVt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        public const string NsXml = "http://www.w3.org/XML/1998/namespace";

        public static XmlDocument LoadXml(string xml)
        {
            XmlDocument doc = new XmlDocument();
            doc.PreserveWhitespace = true;
            doc.XmlResolver = null;
            doc.LoadXml(xml);
            return doc;
        }

        public static XmlNamespaceManager Ns(XmlDocument doc)
        {
            XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("m", NsMain);
            ns.AddNamespace("r", NsRel);
            ns.AddNamespace("pr", NsPkgRel);
            ns.AddNamespace("ct", NsContentTypes);
            ns.AddNamespace("cp", NsCustom);
            ns.AddNamespace("vt", NsVt);
            return ns;
        }

        public static string ReadEntry(ZipArchive zip, string name)
        {
            ZipArchiveEntry e = zip.GetEntry(name);
            if (e == null) return null;
            using (Stream s = e.Open())
            using (StreamReader r = new StreamReader(s, new UTF8Encoding(false), true))
            {
                return r.ReadToEnd();
            }
        }

        public static string ColName(int col)
        {
            if (col <= 0) return "";
            StringBuilder sb = new StringBuilder();
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                col = (col - 1) / 26;
            }
            return sb.ToString();
        }

        public static int ColIndex(string name)
        {
            if (string.IsNullOrEmpty(name)) return 0;
            int col = 0;
            for (int i = 0; i < name.Length; i++)
            {
                char c = char.ToUpperInvariant(name[i]);
                if (c < 'A' || c > 'Z') break;
                col = col * 26 + (c - 'A' + 1);
            }
            return col;
        }

        public static string CellRef(int col, int row)
        {
            return ColName(col) + row.ToString(CultureInfo.InvariantCulture);
        }

        public static string CellColumn(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef)) return "";
            int i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
            return cellRef.Substring(0, i);
        }

        public static int CellRow(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef)) return 0;
            int i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
            int row;
            int.TryParse(cellRef.Substring(i), out row);
            return row;
        }

        public static string Attr(XmlElement e, string name)
        {
            if (e == null || !e.HasAttribute(name)) return "";
            return e.GetAttribute(name);
        }

        public static XmlElement Child(XmlElement e, string localName)
        {
            if (e == null) return null;
            foreach (XmlNode n in e.ChildNodes)
            {
                if (n.NodeType == XmlNodeType.Element && n.LocalName == localName) return (XmlElement)n;
            }
            return null;
        }

        public static string CellText(XmlElement c, List<string> shared)
        {
            if (c == null) return "";
            string t = Attr(c, "t");
            XmlElement v = Child(c, "v");
            if (t == "s")
            {
                int idx;
                if (v != null && int.TryParse(v.InnerText, out idx) && shared != null && idx >= 0 && idx < shared.Count)
                    return shared[idx];
                return "";
            }
            if (t == "inlineStr")
            {
                XmlElement isEl = Child(c, "is");
                if (isEl == null) return "";
                StringBuilder sb = new StringBuilder();
                foreach (XmlNode n in isEl.SelectNodes(".//*[local-name()='t']"))
                    sb.Append(n.InnerText);
                return sb.ToString();
            }
            if (v == null) return "";
            if (t == "b") return v.InnerText == "1" ? "TRUE" : "FALSE";
            return v.InnerText;
        }

        public static decimal CellNumber(XmlElement c, List<string> shared, decimal fallback)
        {
            string text = CellText(c, shared);
            return Format.ParseDecimal(text, fallback);
        }

        public static List<string> ReadShared(ZipArchive zip, out string rawXml)
        {
            rawXml = ReadEntry(zip, "xl/sharedStrings.xml");
            List<string> list = new List<string>();
            if (string.IsNullOrEmpty(rawXml)) return list;
            XmlDocument doc = LoadXml(rawXml);
            XmlNamespaceManager ns = Ns(doc);
            XmlNodeList sis = doc.SelectNodes("/m:sst/m:si", ns);
            if (sis != null)
            {
                foreach (XmlNode si in sis)
                {
                    StringBuilder sb = new StringBuilder();
                    XmlNodeList ts = si.SelectNodes(".//m:t", ns);
                    if (ts != null) foreach (XmlNode t in ts) sb.Append(t.InnerText);
                    list.Add(sb.ToString());
                }
            }
            return list;
        }

        public static string FindSheetPath(ZipArchive zip)
        {
            string wbXml = ReadEntry(zip, "xl/workbook.xml");
            if (string.IsNullOrEmpty(wbXml)) return "xl/worksheets/sheet1.xml";
            XmlDocument wb = LoadXml(wbXml);
            XmlNamespaceManager ns = Ns(wb);
            XmlElement sheet = wb.SelectSingleNode("/m:workbook/m:sheets/m:sheet", ns) as XmlElement;
            if (sheet == null) return "xl/worksheets/sheet1.xml";
            string rid = sheet.GetAttribute("id", NsRel);
            if (string.IsNullOrEmpty(rid)) return "xl/worksheets/sheet1.xml";
            string relXml = ReadEntry(zip, "xl/_rels/workbook.xml.rels");
            if (string.IsNullOrEmpty(relXml)) return "xl/worksheets/sheet1.xml";
            XmlDocument rel = LoadXml(relXml);
            XmlNamespaceManager rns = new XmlNamespaceManager(rel.NameTable);
            rns.AddNamespace("pr", NsPkgRel);
            foreach (XmlElement e in rel.SelectNodes("/pr:Relationships/pr:Relationship", rns))
            {
                if (e.GetAttribute("Id") == rid)
                {
                    string target = e.GetAttribute("Target");
                    if (string.IsNullOrEmpty(target)) break;
                    if (target.StartsWith("/")) return target.TrimStart('/');
                    return NormalizePath("xl/" + target);
                }
            }
            return "xl/worksheets/sheet1.xml";
        }

        public static string NormalizePath(string p)
        {
            if (string.IsNullOrEmpty(p)) return p;
            p = p.Replace('\\', '/');
            string[] parts = p.Split('/');
            List<string> stack = new List<string>();
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i] == "" || parts[i] == ".") continue;
                if (parts[i] == "..")
                {
                    if (stack.Count > 0) stack.RemoveAt(stack.Count - 1);
                }
                else stack.Add(parts[i]);
            }
            return string.Join("/", stack.ToArray());
        }

        public static DateTime? ParseDateTime(string s)
        {
            DateTime d;
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out d))
                return d;
            return null;
        }

        public static string XmlEscape(string s)
        {
            if (s == null) return "";
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                    .Replace("\"", "&quot;").Replace("'", "&apos;");
        }

        public static byte[] XmlBytes(XmlDocument doc)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UTF8Encoding(false);
            settings.Indent = false;
            settings.OmitXmlDeclaration = false;
            using (MemoryStream ms = new MemoryStream())
            {
                using (XmlWriter w = XmlWriter.Create(ms, settings))
                {
                    doc.Save(w);
                }
                return ms.ToArray();
            }
        }

        public static string SafeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "班级";
            char[] invalid = Path.GetInvalidFileNameChars();
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < name.Length; i++)
            {
                bool bad = false;
                for (int j = 0; j < invalid.Length; j++) if (name[i] == invalid[j]) { bad = true; break; }
                sb.Append(bad ? '_' : name[i]);
            }
            string s = sb.ToString().Trim();
            return s.Length == 0 ? "班级" : s;
        }
    }

    /// <summary>解析“方面X事项（分数分）”明细。</summary>
    public static class DetailParser
    {
        private static readonly Regex ScoreRegex = new Regex(
            @"[（(]\s*([+-]?\d+(?:\.\d+)?)\s*分\s*[）)]|[（(]\s*([+-]?\d+\.\d+)\s*[）)]",
            RegexOptions.Compiled);

        private static readonly Regex AspectRegex = new Regex(
            @"^方面\s*([一二三四五六七八九十百零〇0-9]+)",
            RegexOptions.Compiled);

        public static List<RecordItem> Parse(string text, out int maxAspect)
        {
            List<RecordItem> list = new List<RecordItem>();
            maxAspect = 0;
            if (string.IsNullOrEmpty(text)) return list;

            int prevEnd = 0;
            int lastAspect = 0;
            string lastLabel = "";
            MatchCollection matches = ScoreRegex.Matches(text);
            for (int i = 0; i < matches.Count; i++)
            {
                Match m = matches[i];
                string num = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
                decimal delta = Format.ParseDecimal(num, 0m);
                int rawStart = prevEnd;
                int bodyEnd = m.Index;
                int rawEnd = m.Index + m.Length;
                if (rawEnd < rawStart) rawEnd = rawStart;
                string body = text.Substring(rawStart, bodyEnd - rawStart);
                string raw = text.Substring(rawStart, rawEnd - rawStart);

                int aspect = 0;
                string item = body.Trim();
                Match am = AspectRegex.Match(item);
                if (am.Success)
                {
                    aspect = Cn.Parse(am.Groups[1].Value);
                    item = item.Substring(am.Length).Trim();
                }
                else
                {
                    aspect = lastAspect;
                }
                if (aspect > 0)
                {
                    lastAspect = aspect;
                    lastLabel = ClassData.AspectLabel(aspect);
                    if (aspect > maxAspect) maxAspect = aspect;
                }

                RecordItem r = new RecordItem();
                r.AspectIndex = aspect;
                r.AspectLabel = aspect > 0 ? ClassData.AspectLabel(aspect) : "";
                r.Item = item;
                r.Delta = delta;
                r.Raw = raw;
                r.Imported = true;
                list.Add(r);
                prevEnd = rawEnd;
            }

            if (prevEnd < text.Length)
            {
                string tail = text.Substring(prevEnd);
                if (tail.Trim().Length > 0)
                {
                    RecordItem r = new RecordItem();
                    r.AspectIndex = lastAspect;
                    r.AspectLabel = lastAspect > 0 ? lastLabel : "";
                    r.Item = tail.Trim();
                    r.Delta = 0m;
                    r.Raw = tail;
                    r.Imported = true;
                    list.Add(r);
                }
            }

            if (list.Count == 0 && text.Trim().Length > 0)
            {
                RecordItem r = new RecordItem();
                r.AspectIndex = 0;
                r.AspectLabel = "";
                r.Item = text.Trim();
                r.Delta = 0m;
                r.Raw = text;
                r.Imported = true;
                list.Add(r);
            }
            return list;
        }
    }
}
