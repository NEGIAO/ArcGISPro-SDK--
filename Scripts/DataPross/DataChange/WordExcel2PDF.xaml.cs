using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
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

namespace CCTool.Scripts.DataPross.DataChange
{
    /// <summary>
    /// Interaction logic for WordExcel2PDF.xaml
    /// </summary>
    public partial class WordExcel2PDF : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "WordExcel2PDF";

        public WordExcel2PDF()
        {
            InitializeComponent();

            // 初始化其它参数选项
            textExcelFolder.Text = BaseTool.ReadValueFromReg(toolSet, "excelFolder");
            textPDFFolder.Text = BaseTool.ReadValueFromReg(toolSet, "pdfFolder");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "Word、Excel文件批量转PDF";

        private void openExcelFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelFolder.Text = UITool.OpenDialogFolder();
        }

        private void openPDFFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textPDFFolder.Text = UITool.OpenDialogFolder();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string excelFolder = textExcelFolder.Text;
                string pdfFolder = textPDFFolder.Text;

                // 判断参数是否选择完全
                if (excelFolder == "" || pdfFolder == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "excelFolder", excelFolder);
                BaseTool.WriteValueToReg(toolSet, "pdfFolder", pdfFolder);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                // 异步执行
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(excelFolder, pdfFolder);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    // 获取所有excel文件并转换
                    List<string> keys = new List<string>() { ".xls", ".xlsx" };
                    List<string> excelFiles = DirTool.GetAllFilesFromList(excelFolder, keys);
                    foreach (var excelFile in excelFiles)
                    {
                        string fileName = excelFile[(excelFile.LastIndexOf(@"\") + 1)..excelFile.LastIndexOf(".xls")];
                        string pdfFile = $@"{pdfFolder}\{fileName}.pdf";

                        pw.AddMessageMiddle(0, $"      转换文件：{excelFile[(excelFile.LastIndexOf(@"\") + 1)..]}", Brushes.Gray);
                        ExcelTool.ImportToPDF(excelFile, pdfFile);
                    }

                    // 获取所有word文件并转换
                    List<string> keys2 = new List<string>() { ".doc", ".docx" };
                    List<string> wordFiles = DirTool.GetAllFilesFromList(excelFolder, keys2);
                    foreach (var wordFile in wordFiles)
                    {
                        string fileName = wordFile[(wordFile.LastIndexOf(@"\") + 1)..wordFile.LastIndexOf(".doc")];
                        string pdfFile = $@"{pdfFolder}\{fileName}.pdf";

                        pw.AddMessageMiddle(0, $"      转换文件：{wordFile[(wordFile.LastIndexOf(@"\") + 1)..]}", Brushes.Gray);
                        WordTool.ImportToPDF(wordFile, pdfFile);
                    }


                    pw.AddMessageEnd();
                });

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/148134266";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string excelFolder, string pdfFolder)
        {
            List<string> result = new List<string>();

            // 检查路径是否存在
            string result_value = CheckTool.CheckFolderExists(excelFolder);
            if (result_value != "")
            {
                result.Add(result_value);
            }
            string result_value2 = CheckTool.CheckFolderExists(pdfFolder);
            if (result_value2 != "")
            {
                result.Add(result_value2);
            }

            return result;
        }
    }
}
