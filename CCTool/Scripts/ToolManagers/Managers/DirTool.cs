using SharpCompress.Archives;
using SharpCompress.Archives.Rar;
using SharpCompress.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Managers
{
    public class DirTool
    {
        // 获取输入文件夹下的所有文件，带后缀名查询
        public static List<string> GetAllFiles(string folderPath, string keyWord = "no match")
        {
            List<string> filePaths = new List<string>();

            // 获取当前文件夹下的所有文件
            string[] files = Directory.GetFiles(folderPath);
            // 判断是否包含关键字
            if (keyWord == "no match")
            {
                filePaths.AddRange(files);
            }
            else
            {
                foreach (string file in files)
                {
                    // 检查文件名是否包含指定扩展名
                    if (Path.GetExtension(file).Equals(keyWord, StringComparison.OrdinalIgnoreCase))
                    {
                        filePaths.Add(file);
                    }
                }
            }

            // 获取当前文件夹下的所有子文件夹
            string[] subDirectories = Directory.GetDirectories(folderPath);

            // 递归遍历子文件夹下的文件
            foreach (string subDirectory in subDirectories)
            {
                filePaths.AddRange(GetAllFiles(subDirectory, keyWord));
            }

            return filePaths;
        }

        // 获取输入文件夹下的所有文件，带多个后缀名查询
        public static List<string> GetAllFilesFromList(string folderPath, List<string> keyWords = null)
        {
            List<string> filePaths = new List<string>();

            // 获取当前文件夹下的所有文件
            string[] files = Directory.GetFiles(folderPath);
            // 判断是否包含关键字
            if (keyWords == null)
            {
                filePaths.AddRange(files);
            }
            else
            {
                foreach (string file in files)
                {
                    // 标记
                    bool isHave = false;
                    foreach (var key_word in keyWords)
                    {
                        // 检查文件名是否包含指定扩展名
                        if (Path.GetExtension(file).Equals(key_word, StringComparison.OrdinalIgnoreCase))
                        {
                            isHave = true;
                            break;
                        }
                    }
                    // 加入
                    if (isHave)
                    {
                        filePaths.Add(file);
                    }
                }
            }

            // 获取当前文件夹下的所有子文件夹
            string[] subDirectories = Directory.GetDirectories(folderPath);

            // 递归遍历子文件夹下的文件
            foreach (string subDirectory in subDirectories)
            {
                filePaths.AddRange(GetAllFilesFromList(subDirectory, keyWords));
            }

            return filePaths;
        }

        // 获取输入文件夹下的所有GDB文件
        public static List<string> GetAllGDBFilePaths(string folderPath)
        {
            List<string> gdbFilePaths = new List<string>();
            DirectoryInfo directoryInfo = new DirectoryInfo(folderPath);

            // 查找所有GDB数据库文件（.gdb文件夹）
            DirectoryInfo[] gdbDirectories = directoryInfo.GetDirectories("*.gdb", SearchOption.AllDirectories);
            foreach (DirectoryInfo gdbDirectory in gdbDirectories)
            {
                // 获取GDB数据库的路径
                string gdbPath = gdbDirectory.FullName.Replace(@"/", @"\");

                // 添加到列表中
                gdbFilePaths.Add(gdbPath);
            }

            return gdbFilePaths;
        }

        // 从嵌入资源中复制文件
        public static void CopyResourceFile(string resourceName, string filePath)
        {
            // 获取当前程序集的实例
            Assembly assembly = Assembly.GetExecutingAssembly();
            // 从嵌入资源中读取文件
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                // 创建目标文件
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
                {
                    // 将文件从嵌入资源复制到目标文件
                    stream.CopyTo(fileStream);
                }
            }
        }

        // 从嵌入资源中复制压缩包
        public static void CopyResourceRar(string resourceName, string filePath)
        {
            // 获取当前程序集的实例
            Assembly assembly = Assembly.GetExecutingAssembly();
            // 从嵌入资源中读取文件
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                // 创建目标文件
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
                {
                    // 将文件从嵌入资源复制到目标文件
                    stream.CopyTo(fileStream);
                }
            }
            // 解压缩
            using (Stream stream2 = File.OpenRead(filePath))
            {
                using (var reader = RarArchive.Open(stream2))
                {
                    foreach (var entry in reader.Entries)
                    {
                        if (!entry.IsDirectory)
                        {
                            string to_path = filePath[..filePath.LastIndexOf(@"\")];    // 解压位置
                            entry.WriteToDirectory(to_path, new ExtractionOptions()
                            {
                                ExtractFullPath = true,
                                Overwrite = true
                            });
                        }
                    }
                }
            }
            // 删除压缩包
            File.Delete(filePath);
        }

        // 复制文件夹下的所有文件到新的位置
        public static void CopyAllFiles(string sourceDir, string destDir)
        {
            //目标目录不存在则创建
            if (!Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }
            DirectoryInfo sourceDireInfo = new DirectoryInfo(sourceDir);
            List<FileInfo> fileList = new List<FileInfo>();
            GetFileList(sourceDireInfo, fileList); // 获取源文件夹下的所有文件
            List<DirectoryInfo> dirList = new List<DirectoryInfo>();
            GetDirList(sourceDireInfo, dirList); // 获取源文件夹下的所有子文件夹
            // 创建目标文件夹结构
            foreach (DirectoryInfo dir in dirList)
            {
                string sourcePath = dir.FullName;
                string destPath = sourcePath.Replace(sourceDir, destDir); // 替换源文件夹路径为目标文件夹路径
                if (!Directory.Exists(destPath))
                {
                    Directory.CreateDirectory(destPath); // 创建目标文件夹
                }
            }
            // 复制文件到目标文件夹
            foreach (FileInfo fileInfo in fileList)
            {
                string sourceFilePath = fileInfo.FullName;
                string destFilePath = sourceFilePath.Replace(sourceDir, destDir); // 替换源文件夹路径为目标文件夹路径
                File.Copy(sourceFilePath, destFilePath, true); // 复制文件，允许覆盖目标文件
            }
        }

        // 递归获取文件列表
        public static void GetFileList(DirectoryInfo dir, List<FileInfo> fileList)
        {
            fileList.AddRange(dir.GetFiles()); // 添加当前文件夹下的所有文件
            foreach (DirectoryInfo directory in dir.GetDirectories())
            {
                GetFileList(directory, fileList); // 递归获取子文件夹下的文件
            }
        }

        // 递归获取子文件夹列表
        public static void GetDirList(DirectoryInfo dir, List<DirectoryInfo> dirList)
        {
            dirList.AddRange(dir.GetDirectories()); // 添加当前文件夹下的所有子文件夹

            foreach (DirectoryInfo directory in dir.GetDirectories())
            {
                GetDirList(directory, dirList); // 递归获取子文件夹下的子文件夹
            }
        }

        // 创建新文件夹，如果已有，先删再建
        public static void CreateFolder(string path)
        {
            // 创建单元号文件夹
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
            Directory.CreateDirectory(path);
        }

    }
}
