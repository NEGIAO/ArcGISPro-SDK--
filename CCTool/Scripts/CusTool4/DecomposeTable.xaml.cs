using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
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
using ArcGIS.Desktop.Framework.Threading.Tasks;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace CCTool.Scripts.CusTool4
{
    /// <summary>
    /// Interaction logic for DecomposeTable.xaml
    /// </summary>
    public partial class DecomposeTable : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "DecomposeTable";
        public DecomposeTable()
        {
            InitializeComponent();

            // 初始化其它参数选项
            textExcelPath.Text = BaseTool.ReadValueFromReg(toolSet, "excelPath");
            textOutExcelFolder.Text = BaseTool.ReadValueFromReg(toolSet, "outExcelFolder");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "征地一户一表(歌)";

        private void openExcelPathButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.OpenDialogExcel();
        }

        private void openOutExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textOutExcelFolder.Text = UITool.OpenDialogFolder();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string excelPath = textExcelPath.Text;
                string outExcelFolder = textOutExcelFolder.Text;

                // 判断参数是否选择完全
                if (excelPath == "" || outExcelFolder == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "excelPath", excelPath);
                BaseTool.WriteValueToReg(toolSet, "outExcelFolder", outExcelFolder);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                // 异步执行
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(outExcelFolder);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10, $"收集征地人员信息");

                    // 收集信息
                    List<FamilyAtt> fas = new List<FamilyAtt>();

                    // 获取工作薄、工作表
                    string excelFile = ExcelTool.GetPath(excelPath);
                    // 打开工作薄
                    Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                    // 所有工作表
                    var sheets = wb.Worksheets;

                    foreach (Worksheet sheet in sheets) 
                    {
                        // 获取Cells
                        Cells cells = sheet.Cells;
                        // 收集信息
                        FamilyAtt fa = new FamilyAtt();
                        fa.XZQ = sheet.Name;
                        fa.ID = 1;
                        fa.MJ = 0;
                        fa.PeopleCount = 0;
                        fa.PeopleData = new List<List<string>>();

                        // 逐行
                        for (int i = 2; i <= cells.MaxDataRow; i++) 
                        {
                            var idValue = cells[i, 0];
                            string idStr = idValue is null ? "" : idValue.StringValue;

                            int ID = idStr.ToInt();
                            // 如果是当前户
                            if (ID == fa.ID ||  ID == 0)
                            {
                                // 添加信息
                                fa.PeopleCount += 1;
                                fa.MJ += cells[i, 6].DoubleValue;

                                List<string> data = new List<string>();
                                for (int j = 1; j < 12; j++)
                                {
                                    data.Add(cells[i, j]?.StringValue);
                                }
                                fa.PeopleData.Add(data);
                            }
                            // 如果到下一户了,新建一个,重置属性
                            else
                            {
                                // 保存
                                fas.Add(fa);
                                // 重置
                                fa = new FamilyAtt();
                                fa.XZQ = sheet.Name;
                                fa.ID = ID;
                                fa.MJ = 0;
                                fa.PeopleData = new List<List<string>>();

                                // 赋值
                                fa.PeopleCount += 1;
                                fa.MJ += cells[i, 6].DoubleValue;

                                List<string> data = new List<string>();
                                for (int j = 1; j < 12; j++)
                                {
                                    data.Add(cells[i, j]?.StringValue);
                                }
                                fa.PeopleData.Add(data);
                            }
                        }


                    }

                    wb.Dispose();

                    pw.AddMessageMiddle(10, $"输出表一户一表：");
                    // 输出一户一表
                    foreach (var fa in fas)
                    {
                        pw.AddMessageMiddle(0, $"    {fa.XZQ}_{fa.ID}_{fa.PeopleData[0][0]}", Brushes.Gray);

                        // 复制界址点Excel表
                        string path = $@"{outExcelFolder}\{fa.XZQ}_{fa.ID}_{fa.PeopleData[0][0]}.xlsx";
                        DirTool.CopyResourceFile(@"CCTool.Data.Excel.杂七杂八.一户一表.xlsx", path);

                        // 获取工作薄、工作表
                        string excelFile2 = ExcelTool.GetPath(path);
                        int sheetIndex2 = ExcelTool.GetSheetIndex(path);
                        // 打开工作薄
                        Workbook wb2 = ExcelTool.OpenWorkbook(excelFile2);
                        // 打开工作表
                        Worksheet ws2 = wb2.Worksheets[sheetIndex2];
                        // 获取Cells
                        Cells cells = ws2.Cells;

                        // 写入
                        string title = cells["A2"].StringValue;
                        cells["A2"].Value = title.Replace("宜州区     乡镇      村（社区）", fa.XZQ);

                        string area = cells["A3"].StringValue;
                        cells["A3"].Value = area.Replace("0.4240", fa.MJ.RoundWithFill(4));

                        cells["M6"].Value = fa.PeopleCount;

                        for (int j = 0; j < fa.PeopleCount; j++)
                        {

                            for (int k = 0; k < 11; k++)
                            {
                                cells[j + 5, k + 1].Value = fa.PeopleData[j][k].ToString();
                            }
                        }
                        // 保存
                        wb2.Save(excelFile2);
                        wb2.Dispose();
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


        private List<string> CheckData(string outExcelFolder)
        {
            List<string> result = new List<string>();

            // 检查路径是否存在
            string result_value = CheckTool.CheckFolderExists(outExcelFolder);
            if (result_value != "")
            {
                result.Add(result_value);
            }

            return result;
        }
    }
}

class FamilyAtt
{
    public string XZQ { get; set; }
    public int ID {  get; set; }
    public double MJ { get; set; }

    public int PeopleCount { get;set; }

    public List<List<string>> PeopleData { get; set; }

}
