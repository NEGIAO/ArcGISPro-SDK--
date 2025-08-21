using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Row = ArcGIS.Core.Data.Row;

namespace CCTool.Scripts.GHApp.QT
{
    /// <summary>
    /// Interaction logic for LandTransfer.xaml
    /// </summary>
    public partial class LandTransfer : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "LandTransfer";
        public LandTransfer()
        {
            InitializeComponent();

            // 初始化combox
            combox_unit.Items.Add("平方米");
            combox_unit.Items.Add("公顷");
            combox_unit.Items.Add("平方公里");
            combox_unit.Items.Add("亩");
            combox_unit.SelectedIndex = 0;

            combox_digit.Items.Add("1");
            combox_digit.Items.Add("2");
            combox_digit.Items.Add("3");
            combox_digit.Items.Add("4");
            combox_digit.Items.Add("5");
            combox_digit.Items.Add("6");
            combox_digit.SelectedIndex = 1;

            // 初始化其它参数选项
            textExcelPath.Text = BaseTool.ReadValueFromReg(toolSet, "excelPath");
        }
        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "土地利用转移矩阵";

        private void combox_yd1_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_yd1);
        }

        private void combox_yd2_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_yd2);
        }

        private void combox_field1_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_yd1.ComboxText(), combox_field1);
        }

        private void combox_field2_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_yd2.ComboxText(), combox_field2);
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.SaveDialogExcel();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string yd1 = combox_yd1.ComboxText();
                string yd2 = combox_yd2.ComboxText();
                string field1 = combox_field1.ComboxText();
                string field2 = combox_field2.ComboxText();

                string unit = combox_unit.Text;
                int digit = combox_digit.Text.ToInt();

                string excelPath = textExcelPath.Text;
                // 默认数据库位置
                var defGDB = Project.Current.DefaultGeodatabasePath;
                // 工程默认文件夹位置
                string defFolder = Project.Current.HomeFolderPath;

                // 单位系数设置
                double unit_xs = unit switch
                {
                    "平方米" => 1,
                    "公顷" => 10000,
                    "平方公里" => 1000000,
                    "亩" => 666.66667,
                    _ => 1,
                };

                // 判断参数是否选择完全
                if (yd1 == "" || yd2 == "" || field1 == "" || field2 == "" || excelPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }
                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "excelPath", excelPath);


                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(yd1, yd2, field1, field2);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }
                    // 复制嵌入资源中的Excel文件
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.杂七杂八.土地利用转移矩阵.xlsx", excelPath);

                    pw.AddMessageMiddle(10, "数据预处理");
                    string bj1 = "BJFD1";
                    string bj2 = "BJFD2";
                    Arcpy.AddField(yd1, bj1, "TEXT");
                    Arcpy.AddField(yd2, bj2, "TEXT");
                    Arcpy.CalculateField(yd1, bj1, $"!{field1}!");
                    Arcpy.CalculateField(yd2, bj2, $"!{field2}!");

                    pw.AddMessageMiddle(20, "相交");
                    string intersect = @$"{defGDB}\intersect";
                    Arcpy.Intersect(new List<string>() { yd1, yd2 }, intersect);

                    pw.AddMessageMiddle(10, "汇总统计");
                    string statistics = @$"{defGDB}\statistics";
                    Arcpy.Statistics(intersect, statistics, "Shape_Area SUM", $"{bj1};{bj2}");

                    pw.AddMessageMiddle(10, "数据透视");
                    string pivotTable = @$"{defGDB}\pivotTable";
                    Arcpy.PivotTable(statistics, bj1, bj2, "SUM_Shape_Area", pivotTable);

                    pw.AddMessageMiddle(20, "写入Excel");
                    // 获取工作薄、工作表
                    string excelFile = ExcelTool.GetPath(excelPath);
                    int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
                    // 打开工作薄
                    Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                    // 打开工作表
                    Worksheet sheet = wb.Worksheets[sheetIndex];
                    Cells cells = sheet.Cells;
                    int rowIndex = 2;

                    // 字段表
                    List<Field> fds = GisTool.GetFieldsFromTarget(pivotTable);
                    // 取字段别名成列表，避免数字型字段
                    List<string> fields = fds.Select(s => s.AliasName).ToList();
                    // 移除不相干的字段，并排序
                    fields.Remove("OBJECTID");
                    fields.Remove(bj1);
                    fields.Sort();

                    // 读取属性表
                    Table table = pivotTable.TargetTable();
                    using var cursor = table.Search();
                    while (cursor.MoveNext())
                    {
                        Row row = cursor.Current;

                        // 第一次，先复制列, 填写列名
                        if (rowIndex == 2)
                        {
                            for (int i = 0; i < fields.Count; i++)
                            {
                                cells.CopyColumn(cells, 1, i + 1);
                                cells[0, i + 1].Value = fields[i];
                            }
                        }

                        // 复制行
                        cells.CopyRow(cells, 1, rowIndex);
                        // 标记字段1
                        cells[rowIndex, 0].Value = row[bj1]?.ToString();

                        // 面积值
                        for (int i = 0; i < fields.Count; i++)
                        {
                            string fieldValue = row[fields[i]]?.ToString();

                            if (fieldValue is not null)
                            {
                                // 填写
                                cells[rowIndex, i + 1].Value = Math.Round(fieldValue.ToDouble() / unit_xs, digit);
                            }

                        }

                        rowIndex++;
                    }

                    // 删除示例行
                    cells.DeleteRow(1);

                    // 保存
                    wb.Save(excelFile);
                    wb.Dispose();

                    // 删除中间数据
                    Arcpy.Delect(intersect);
                    Arcpy.Delect(statistics);
                    Arcpy.Delect(pivotTable);

                    Arcpy.DeleteField(yd1, bj1);
                    Arcpy.DeleteField(yd2, bj2);

                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://mp.weixin.qq.com/s/b4vlFDJsHi0oAgHKQLyaYw";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string yd1, string yd2, string field1, string field2)
        {
            List<string> result = new List<string>();
            // 检查字段值是否为空
            string fieldEmptyResult1 = CheckTool.CheckFieldValueSpace(yd1, field1);
            if (fieldEmptyResult1 != "")
            {
                result.Add(fieldEmptyResult1);
            }
            string fieldEmptyResult2 = CheckTool.CheckFieldValueSpace(yd2, field2);
            if (fieldEmptyResult2 != "")
            {
                result.Add(fieldEmptyResult2);
            }

            return result;
        }
    }
}
