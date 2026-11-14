using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace CaoXing
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += delegate(object sender, ThreadExceptionEventArgs e)
            {
                ReportError(e.Exception);
            };
            AppDomain.CurrentDomain.UnhandledException += delegate(object sender, UnhandledExceptionEventArgs e)
            {
                ReportError(e.ExceptionObject as Exception);
            };

            string open = "";
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "--open" && i + 1 < args.Length)
                {
                    open = args[i + 1];
                    break;
                }
                if (!args[i].StartsWith("-") && open.Length == 0)
                {
                    open = args[i];
                }
            }
            Application.Run(new MainForm(open));
        }

        private static void ReportError(Exception ex)
        {
            string message = ex == null ? "未知错误" : ex.ToString();
            try
            {
                string path = Path.Combine(Application.StartupPath, "错误日志.txt");
                File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\r\n" + message + "\r\n\r\n", new UTF8Encoding(true));
            }
            catch { }
            try
            {
                MessageBox.Show("程序遇到一个错误：\r\n\r\n" + (ex == null ? "未知错误" : ex.Message) +
                    "\r\n\r\n详细信息已写入程序目录的“错误日志.txt”。",
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch { }
        }
    }
}
