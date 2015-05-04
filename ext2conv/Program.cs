using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Collections;

namespace ext2conv
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] picExts = {".bmp", ".png", ".jpeg", ".jpg", ".gif", ".tiff"};
            string[] officeExts = { ".doc", ".docx", ".ppt", ".pptx" };

            Hashtable PathHash = new Hashtable();
            PathHash["convert"]   = "C:\\Program Files\\ImageMagick-6.9.0-Q16\\convert.exe";
            PathHash["soffice"]   = "C:\\Program Files (x86)\\LibreOffice 4\\program\\soffice.exe";
            PathHash["tmp"] = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + @"\ext2conv\";

            while (true) {
                var watcher = new FileSystemWatcher();
                watcher.Path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\";
                watcher.Filter = "";
                watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName;
                watcher.IncludeSubdirectories = false;
                var changed = watcher.WaitForChanged(WatcherChangeTypes.All);

                if (changed.ChangeType != WatcherChangeTypes.Renamed) continue;
                var oldExt = Path.GetExtension(changed.OldName);
                var ext = Path.GetExtension(changed.Name);

                Console.Write(changed.Name + "\n");

                if (oldExt.Equals(ext)) continue;
                if (Array.IndexOf(picExts, oldExt) != -1 && Array.IndexOf(picExts, ext) != -1)
                {
                    var psi = new ProcessStartInfo();
                    psi.FileName = (string)PathHash["convert"];
                    psi.Arguments = string.Join(" ", new string[]{
                       "\"" + watcher.Path + changed.Name + "\"",
                       "\"" + watcher.Path + changed.Name + "\""
                    });
                    psi.WindowStyle = ProcessWindowStyle.Hidden;
                    Process.Start(psi);
                }

                if (Array.IndexOf(officeExts, oldExt) != -1 && ext.Equals(".pdf"))
                {
                    string documentsPath = (string)PathHash["tmp"];
                    if (Directory.Exists(documentsPath)) Directory.Delete(documentsPath, true);
                    Directory.CreateDirectory(documentsPath);

                    File.Move(watcher.Path + changed.Name, documentsPath + changed.Name);
                    var psi = new ProcessStartInfo();
                    psi.WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    psi.FileName = (string)PathHash["soffice"];
                    psi.Arguments = string.Join(" ", new string[]{
                        "--headless",
                        "--convert-to pdf",
                        "\"" + documentsPath + changed.Name + "\""
                    });
                    psi.WindowStyle = ProcessWindowStyle.Hidden;
                    Process.Start(psi);

                    while (true)
                    {
                        if (File.Exists(watcher.Path + changed.Name)) break;
                        System.Threading.Thread.Sleep(100);
                    }

                    Directory.Delete(documentsPath, true);
                }

                if (!oldExt.Equals(".zip") && ext.Equals(".zip"))
                {
                    string documentsPath = (string)PathHash["tmp"];
                    if (Directory.Exists(documentsPath)) Directory.Delete(documentsPath, true);
                    Directory.CreateDirectory(documentsPath);

                    if (File.Exists(watcher.Path + changed.Name))
                      File.Move(watcher.Path + changed.Name, documentsPath + changed.OldName);
                    else if (Directory.Exists(watcher.Path + changed.Name))
                      Directory.Move(watcher.Path + changed.Name, documentsPath + changed.OldName);

                    ZipFile.CreateFromDirectory(documentsPath, watcher.Path + changed.Name);

                    Directory.Delete(documentsPath, true);
                }
            }
        }
    }
}
