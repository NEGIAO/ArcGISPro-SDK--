using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Managers;
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

namespace CCTool.Scripts.CusTool2
{
    /// <summary>
    /// Interaction logic for WordReplace2.xaml
    /// </summary>
    public partial class WordReplace2 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public WordReplace2()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "Excel_Word文本批量替换(冬至)";

        private void openWordButton_Click(object sender, RoutedEventArgs e)
        {
            string wordPath = UITool.OpenDialogFolder();
            textWordPath.Text = wordPath;
            string path = wordPath[..wordPath.LastIndexOf(@"\")];
            textOutPath.Text = path + @"\输出报告";
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.OpenDialogExcel();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string wordPath = textWordPath.Text;
                string excelPath = textExcelPath.Text;
                string outPath = textOutPath.Text;

                // 如果输出路径不存在，就创建一个
                if (!Directory.Exists(outPath))
                {
                    Directory.CreateDirectory(outPath);
                }

                // 判断参数是否选择完全
                if (wordPath == "" || excelPath == "" || outPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("获取指标");
                    // 获取指标
                    List<Dictionary<string, string>> list = ExcelTool.GetDictListFromExcelCol(excelPath);

                    // 获取所有模板
                    List<string> allFiles =DirTool.GetAllFilesFromList(wordPath, new List<string>() { ".doc", ".docx" });

                    foreach (var dict in list)
                    {
                        string bmExcel = dict["模板编号"] ?? "";      // 模板编号
                        string name = dict["项目名称"] ?? "";    // 项目名称

                        pw.AddMessageMiddle(20, $"写入指标_{name}_{bmExcel}");
                        // 复制word
                        foreach (var file in allFiles)
                        {
                            // 获取word中的模板编号
                            string fileName = file[(file.LastIndexOf(@"\") + 1)..];
                            string bmWord = fileName[(fileName.IndexOf("【") + 1)..fileName.IndexOf("】")];

                            if (bmExcel == bmWord)
                            {
                                string outFile = @$"{outPath}\{name}_{bmExcel}.doc";
                                // 复制模板
                                File.Copy(file, outFile, true);
                                foreach (var item in dict)
                                {
                                    WordTool.WordRepalceText(outFile, @"{" + item.Key + @"}", item.Value);
                                }
                            }
                        }
                    }
                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void openOutButton_Click(object sender, RoutedEventArgs e)
        {
            textOutPath.Text = UITool.OpenDialogFolder();
        }
    }
}
