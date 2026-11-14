using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace CaoXing
{
    public static class SaveService
    {
        public static string SuggestFileName(ClassData data)
        {
            string name = string.IsNullOrWhiteSpace(data.ClassName) ? "班级" : data.ClassName.Trim();
            return XlsxCommon.SafeFileName(name) + "操行分汇总_" + DateTime.Now.ToString("yyyy-MM") + ".xlsx";
        }

        /// <summary>
        /// 保存数据。若数据来自模板，则在模板副本上修改；否则生成新表。
        /// 返回备份文件路径（没有备份则为空字符串）。
        /// </summary>
        public static string Save(ClassData data, string targetPath, string templateSource, bool rosterDirty, bool useTemplate)
        {
            if (string.IsNullOrEmpty(targetPath)) throw new XlsxException("没有指定保存路径。");
            string dir = Path.GetDirectoryName(Path.GetFullPath(targetPath));
            if (string.IsNullOrEmpty(dir)) throw new XlsxException("保存路径无效。");
            Directory.CreateDirectory(dir);
            string temp = Path.Combine(dir, "." + Path.GetFileName(targetPath) + ".tmp-" + Guid.NewGuid().ToString("N") + ".xlsx");

            try
            {
                if (useTemplate && data.FromTemplate && !string.IsNullOrEmpty(templateSource) && File.Exists(templateSource))
                {
                    XlsxTemplateWriter.Write(data, templateSource, temp, rosterDirty);
                }
                else
                {
                    XlsxCreator.Create(data, temp);
                }

                string backup = "";
                if (File.Exists(targetPath))
                {
                    backup = Backup(targetPath);
                    try
                    {
                        File.Replace(temp, targetPath, null, true);
                    }
                    catch
                    {
                        File.Copy(temp, targetPath, true);
                    }
                }
                else
                {
                    File.Move(temp, targetPath);
                }
                return backup;
            }
            catch (IOException ex)
            {
                throw new XlsxException("保存失败：文件可能正被 Excel / WPS 打开。请关闭该表格后重试。\r\n\r\n" + ex.Message, ex);
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new XlsxException("保存失败：没有写入权限，或文件为只读。请换一个位置保存。\r\n\r\n" + ex.Message, ex);
            }
            finally
            {
                try { if (File.Exists(temp)) File.Delete(temp); } catch { }
            }
        }

        public static string Backup(string targetPath)
        {
            try
            {
                string dir = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(targetPath)), "表格备份");
                Directory.CreateDirectory(dir);
                string baseName = Path.GetFileNameWithoutExtension(targetPath);
                string ext = Path.GetExtension(targetPath);
                string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                string backup = Path.Combine(dir, baseName + "_备份_" + stamp + ext);
                File.Copy(targetPath, backup, true);
                TrimBackups(dir, baseName + "_备份_");
                return backup;
            }
            catch
            {
                return "";
            }
        }

        private static void TrimBackups(string dir, string prefix)
        {
            try
            {
                List<FileInfo> files = new List<FileInfo>();
                foreach (string f in Directory.GetFiles(dir))
                {
                    FileInfo fi = new FileInfo(f);
                    if (fi.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) files.Add(fi);
                }
                files.Sort(delegate(FileInfo a, FileInfo b) { return b.LastWriteTime.CompareTo(a.LastWriteTime); });
                for (int i = 40; i < files.Count; i++)
                {
                    try { files[i].Delete(); } catch { }
                }
            }
            catch { }
        }
    }
}

