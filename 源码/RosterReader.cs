using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CaoXing
{
    public class RosterResult
    {
        public List<StudentSeed> Students = new List<StudentSeed>();
        public string Detail = "";
    }

    /// <summary>从 xlsx / csv / txt / 粘贴文本中读取“学号 + 姓名”名单。</summary>
    public static class RosterReader
    {
        public static RosterResult ReadFile(string path)
        {
            if (!File.Exists(path)) throw new XlsxException("名单文件不存在：" + path);
            string ext = Path.GetExtension(path).ToLowerInvariant();
            List<string[]> rows;
            if (ext == ".xlsx" || ext == ".xlsm")
            {
                rows = XlsxReader.ReadGrid(path);
            }
            else if (ext == ".csv" || ext == ".txt")
            {
                string text = ReadTextAuto(path);
                rows = ParseText(text);
            }
            else if (ext == ".xls")
            {
                throw new XlsxException("不支持旧版 .xls 文件。请在 Excel / WPS 中另存为 .xlsx 或 .csv 后再导入。");
            }
            else
            {
                throw new XlsxException("不支持的文件类型：" + ext + "。请选择 .xlsx、.csv 或 .txt。");
            }
            return BuildResult(rows);
        }

        public static RosterResult ReadPaste(string text)
        {
            return BuildResult(ParseText(text ?? ""));
        }

        public static List<string[]> ParseText(string text)
        {
            if (string.IsNullOrEmpty(text)) return new List<string[]>();
            text = text.Replace("\r\n", "\n").Replace('\r', '\n');
            char delim = DetectDelimiter(text);
            if (delim == '\0')
            {
                // 空白分隔
                List<string[]> result = new List<string[]>();
                string[] lines = text.Split('\n');
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    if (line.Length == 0) continue;
                    string[] parts = line.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    result.Add(parts);
                }
                return result;
            }
            return ParseDelimited(text, delim);
        }

        private static char DetectDelimiter(string text)
        {
            string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            int checkedLines = 0;
            foreach (string line in lines)
            {
                if (line.Trim().Length == 0) continue;
                if (line.IndexOf('\t') >= 0) return '\t';
                if (line.IndexOf(',') >= 0) return ',';
                if (line.IndexOf(';') >= 0) return ';';
                checkedLines++;
                if (checkedLines >= 3) break;
            }
            return '\0';
        }

        private static List<string[]> ParseDelimited(string text, char delim)
        {
            List<string[]> rows = new List<string[]>();
            List<string> fields = new List<string>();
            StringBuilder field = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '"')
                        {
                            field.Append('"');
                            i++;
                        }
                        else inQuotes = false;
                    }
                    else field.Append(c);
                }
                else
                {
                    if (c == '"') inQuotes = true;
                    else if (c == delim)
                    {
                        fields.Add(field.ToString());
                        field.Length = 0;
                    }
                    else if (c == '\n')
                    {
                        fields.Add(field.ToString());
                        field.Length = 0;
                        rows.Add(fields.ToArray());
                        fields.Clear();
                    }
                    else if (c == '\r') { }
                    else field.Append(c);
                }
            }
            if (field.Length > 0 || fields.Count > 0)
            {
                fields.Add(field.ToString());
                rows.Add(fields.ToArray());
            }
            return rows;
        }

        private static RosterResult BuildResult(List<string[]> rows)
        {
            RosterResult result = new RosterResult();
            if (rows == null || rows.Count == 0)
                throw new XlsxException("没有读到任何数据。");

            int noCol = -1;
            int nameCol = -1;
            int headerIndex = -1;
            int scan = Math.Min(rows.Count, 20);
            for (int i = 0; i < scan; i++)
            {
                string[] row = rows[i];
                int a = FindHeader(row, "学号");
                int b = FindHeader(row, "姓名");
                if (a >= 0 && b >= 0)
                {
                    noCol = a;
                    nameCol = b;
                    headerIndex = i;
                    break;
                }
            }
            if (noCol < 0 || nameCol < 0)
            {
                noCol = 0;
                nameCol = 1;
                headerIndex = -1;
            }

            int start = headerIndex + 1;
            for (int i = start; i < rows.Count; i++)
            {
                string[] row = rows[i];
                if (row == null) continue;
                string no = noCol < row.Length ? row[noCol].Trim() : "";
                string name = nameCol < row.Length ? row[nameCol].Trim() : "";
                if (no.Length == 0 && name.Length == 0) continue;
                if (no == "学号" && name == "姓名") continue;
                if (no.Length == 0) throw new XlsxException(string.Format("第 {0} 条数据缺少学号。", result.Students.Count + 1));
                if (name.Length == 0) throw new XlsxException(string.Format("第 {0} 条数据（学号 {1}）缺少姓名。", result.Students.Count + 1, no));
                if (no.Length > 40) throw new XlsxException(string.Format("学号“{0}”过长，请检查列是否对应正确。", no));
                if (name.Length > 40) throw new XlsxException(string.Format("姓名“{0}”过长，请检查列是否对应正确。", name));
                result.Students.Add(new StudentSeed(no, name));
            }
            if (result.Students.Count == 0) throw new XlsxException("没有读到有效的学生记录。");
            result.Detail = string.Format("共 {0} 人{1}", result.Students.Count,
                headerIndex >= 0 ? "（已识别表头：学号 / 姓名）" : "（未发现表头，按前两列读取）");
            return result;
        }

        private static int FindHeader(string[] row, string key)
        {
            if (row == null) return -1;
            for (int i = 0; i < row.Length; i++)
            {
                string s = row[i].Trim();
                if (s == key || s == key + "：" || s.StartsWith(key)) return i;
            }
            return -1;
        }

        private static string ReadTextAuto(string path)
        {
            byte[] bytes = File.ReadAllBytes(path);
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
                return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
            try
            {
                return new UTF8Encoding(false, true).GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                try { return Encoding.GetEncoding(936).GetString(bytes); }
                catch { return Encoding.Default.GetString(bytes); }
            }
        }
    }
}
