using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Managers;
using SharpCompress.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Path = System.IO.Path;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for UpdataFileName.xaml
    /// </summary>
    public partial class UpdataFileName : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public UpdataFileName()
        {
            InitializeComponent();

            textExcelPath.Text = BaseTool.ReadValueFromReg("UpdataFileName", "folderPath");
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.OpenDialogFolder();
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/144550873";
            UITool.Link2Web(url);
        }

        private void btn_go_Click(object sender, RoutedEventArgs e)
        {
           
            // 获取参数
            string folderPath = textExcelPath.Text;

            BaseTool.WriteValueToReg("UpdataFileName", "folderPath", folderPath);

            // 判断参数是否选择完全
            if (folderPath == "" )
            {
                MessageBox.Show("有必选参数为空！！！");
                return;
            }


            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("指定的文件夹不存在！");
                return;
            }

            
            try
            {
                // 计数
                int count = 0;
                // 对照列表
                Dictionary<string, string> dict = new Dictionary<string, string>();

                // 获取所有子文件夹
                string[] subFolders = Directory.GetDirectories(folderPath);

                foreach (var subFolder in subFolders)
                {

                    // 获取当前子文件夹中的所有 .jpg 文件
                    string[] jpgFiles = Directory.GetFiles(subFolder);

                    for (int i = 0; i < jpgFiles.Length; i++)
                    {
                        string oldFilePath = jpgFiles[i];

                        // 获取文件的目录路径
                        string directoryPath = Path.GetDirectoryName(oldFilePath);
                        // 获取目录路径的上一层级目录名称
                        string parentFolderName = Path.GetFileName(directoryPath);
                        // 文件序号
                        string ID = (i + 1).ToString().PadLeft(3, '0');
                        // 文件格式
                        string fileType = oldFilePath[(oldFilePath.LastIndexOf('.') + 1)..];

                        string newFileName = $"{parentFolderName}Z{ID}.{fileType}"; // 避免文件重名，添加索引后缀
                        string newFilePath = Path.Combine(subFolder, newFileName);

                        File.Move(oldFilePath, newFilePath);
                        // 加入集合
                        dict.Add(newFileName, parentFolderName);

                        count++;
                    }

                    // 生成对照表
                    // 复制嵌入资源中的Excel文件
                    string excelPath = $@"{folderPath}\文件对照表.xlsx";
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.文件对照表2.xlsx", excelPath);

                    // 获取工作薄、工作表
                    string excelFile = ExcelTool.GetPath(excelPath);
                    int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
                    // 打开工作薄
                    Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                    // 打开工作表
                    Worksheet sheet = wb.Worksheets[sheetIndex];

                    Cells cells= sheet.Cells;

                    // 逐行处理
                    int index = 1;
                    foreach (var item in dict)
                    {
                        string folderName = item.Value;
                        string fileName = item.Key;
  
                        cells[index, 0].Value = folderName;   // 赋值
                        cells[index, 1].Value = fileName;   

                        index++;
                    }

                    // 保存
                    wb.Save(excelFile);
                    wb.Dispose();

                }

                Close();
                MessageBox.Show($"文件重命名成功，共计{count}个文件。");
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message+ee.StackTrace);
                return;
            }

        }
    }
}
