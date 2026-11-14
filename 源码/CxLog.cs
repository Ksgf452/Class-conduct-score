using System;
using System.IO;
using System.Text;

namespace CaoXing
{
    /// <summary>按月生成并追加操行日志。</summary>
    public static class CxLog
    {
        public static string MonthFileName(DateTime time)
        {
            return "操行日志-" + time.ToString("yyyy-MM") + ".log";
        }

        public static string ResolveDirectory(string workbookPath, string appDir)
        {
            if (!string.IsNullOrEmpty(workbookPath))
            {
                string dir = Path.GetDirectoryName(Path.GetFullPath(workbookPath));
                if (!string.IsNullOrEmpty(dir)) return Path.Combine(dir, "操行日志");
            }
            return Path.Combine(appDir, "操行日志");
        }

        /// <summary>写一条日志。fields 为“字段=值”形式，异常时静默失败，不打断主流程。</summary>
        public static string Write(string logDir, string action, params string[] fields)
        {
            DateTime now = DateTime.Now;
            return WriteAt(logDir, now, action, fields);
        }

        public static string WriteAt(string logDir, DateTime now, string action, params string[] fields)
        {
            try
            {
                if (string.IsNullOrEmpty(logDir)) return "";
                Directory.CreateDirectory(logDir);
                string path = Path.Combine(logDir, MonthFileName(now));
                bool isNew = !File.Exists(path);
                StringBuilder sb = new StringBuilder();
                if (isNew)
                {
                    sb.Append("# 班级操行记录日志 | ").Append(now.ToString("yyyy年MM月")).Append("\r\n");
                    sb.Append("# 本文件由“班级操行分管理系统”自动生成，请勿用 Excel 打开后另存，以免破坏编码。\r\n");
                    sb.Append("# 字段顺序：时间 | 操作 | 详情...\r\n");
                    sb.Append("\r\n");
                }
                sb.Append(now.ToString("yyyy-MM-dd HH:mm:ss")).Append(" | ").Append(action);
                if (fields != null)
                {
                    for (int i = 0; i < fields.Length; i++)
                    {
                        if (fields[i] == null) continue;
                        sb.Append(" | ").Append(fields[i].Replace("\r", " ").Replace("\n", " "));
                    }
                }
                sb.Append("\r\n");
                File.AppendAllText(path, sb.ToString(), new UTF8Encoding(true));
                return path;
            }
            catch
            {
                return "";
            }
        }
    }
}
