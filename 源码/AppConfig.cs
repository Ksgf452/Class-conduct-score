using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CaoXing
{
    /// <summary>简单的本地设置（记住上次目录、窗口大小）。</summary>
    internal class AppConfig
    {
        public string LastFolder = "";
        public int Width = 1300;
        public int Height = 830;
        public bool Maximized = false;

        private static string FilePath
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.ini"); }
        }

        public static AppConfig Load()
        {
            AppConfig c = new AppConfig();
            try
            {
                if (!File.Exists(FilePath)) return c;
                foreach (string raw in File.ReadAllLines(FilePath, Encoding.UTF8))
                {
                    string line = raw.Trim();
                    if (line.Length == 0 || line.StartsWith("#")) continue;
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string key = line.Substring(0, eq).Trim();
                    string value = line.Substring(eq + 1).Trim();
                    if (key == "LastFolder") c.LastFolder = value;
                    else if (key == "Width") { int.TryParse(value, out c.Width); }
                    else if (key == "Height") { int.TryParse(value, out c.Height); }
                    else if (key == "Maximized") c.Maximized = value == "1";
                }
            }
            catch { }
            if (c.Width < 1120) c.Width = 1300;
            if (c.Height < 720) c.Height = 830;
            return c;
        }

        public void Save()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("# 班级操行分管理系统设置");
                sb.AppendLine("LastFolder=" + (LastFolder ?? ""));
                sb.AppendLine("Width=" + Width);
                sb.AppendLine("Height=" + Height);
                sb.AppendLine("Maximized=" + (Maximized ? "1" : "0"));
                File.WriteAllText(FilePath, sb.ToString(), new UTF8Encoding(false));
            }
            catch { }
        }
    }
}
