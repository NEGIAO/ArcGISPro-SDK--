using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Row = ArcGIS.Core.Data.Row;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for SDStatisticPlus.xaml
    /// </summary>
    public partial class SDStatisticPlus : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "SDStatisticPlus";

        public SDStatisticPlus()
        {
            InitializeComponent();

            // 初始化combox
            combox_mjType.Items.Add("平方米");
            combox_mjType.Items.Add("公顷");
            combox_mjType.Items.Add("平方公里");
            combox_mjType.Items.Add("亩");
            combox_mjType.SelectedIndex = 3;


            // 初始化其它参数选项
            textExcelPath.Text = BaseTool.ReadValueFromReg(toolSet, "excelPath");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "图斑占三调用地统计表(弓)";


        // 运行
        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                string defGDB = Project.Current.DefaultGeodatabasePath;
                string defFolder = Project.Current.HomeFolderPath;

                // 获取参数
                string sd = combox_sd.ComboxText();
                string dk = combox_dk.ComboxText();
                string field = combox_field.ComboxText();

                string excelPath = textExcelPath.Text;

                string unit = combox_mjType.Text;

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
                if (sd == "" || dk == "" || field == "" || excelPath == "")
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
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.用地占用图斑统计表.xlsx", excelPath);

                    pw.AddMessageStart("相交并处理相关字段");
                    // 为避免字段同名，新建一个地块名称字段
                    string dkField = "dkField";
                    Arcpy.AddField(dk, dkField, "TEXT");
                    Arcpy.CalculateField(dk, dkField, $"!{field}!");
                    string mjField = "mjField";
                    // 计算图斑面积
                    Arcpy.AddField(dk, mjField, "DOUBLE");
                    Arcpy.CalculateField(dk, mjField, $"!shape!.area/{unit_xs}");

                    // 相交
                    List<string> fcs = new List<string>() { sd, dk };
                    string intersectResult = $@"{defGDB}\intersectResult";
                    Arcpy.Intersect(fcs, intersectResult);

                    // 地块名称列表
                    Dictionary<string, double> dict_dk = GisTool.GetDictFromPathDouble(dk, dkField, mjField);

                    pw.AddMessageMiddle(20, "写入Excel表");
                    // 获取工作薄、工作表
                    string excelFile = ExcelTool.GetPath(excelPath);
                    int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
                    // 打开工作薄
                    Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                    // 打开工作表
                    Worksheet sheet = wb.Worksheets[sheetIndex];
                    // 获取cell
                    Cells cells = sheet.Cells;

                    // 计数
                    int index = 1;

                    // 按地块名称分开统计
                    foreach (var dk in dict_dk)
                    {
                        // 地块名称和地块面积
                        string dkName = dk.Key;
                        double dkArea = dk.Value;

                        // 按地块名筛选
                        var queryFilter = new QueryFilter();
                        queryFilter.WhereClause = $"{dkField} = '{dkName}'";

                        Table table = intersectResult.TargetTable();
                        using RowCursor rowCursor = table.Search(queryFilter);

                        while (rowCursor.MoveNext())
                        {
                            // 复制行
                            if (index > 1)
                            {
                                cells.CopyRow(cells, 1, index);
                            }

                            using Row row = rowCursor.Current;
                            // 获取参数
                            string TBBH = row["TBBH"]?.ToString();
                            string DLBM = row["DLBM"]?.ToString();
                            string DLMC = row["DLMC"]?.ToString();

                            double MJ = double.Parse(row["shape_area"]?.ToString()) / unit_xs;

                            // 写入参数
                            cells[index, 0].Value = dkName;
                            cells[index, 1].Value = dkArea;
                            cells[index, 2].Value = MJ;
                            cells[index, 3].Value = TBBH;
                            cells[index, 4].Value = DLMC;
                            cells[index, 5].Value = DLBM;
                            index++;
                        }
                    }


                    // 保存
                    wb.Save(excelFile);
                    wb.Dispose();

                    // 合并相同格
                    ExcelTool.MergeSameCol(excelPath, 0);
                    ExcelTool.MergeSameCol(excelPath, 1);

                    // 删除中间数据
                    Arcpy.Delect(intersectResult);
                    // 删除字段
                    Arcpy.DeleteField(dk, dkField);
                    Arcpy.DeleteField(dk, mjField);

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
            string url = "https://blog.csdn.net/xcc34452366/article/details/146005717";
            UITool.Link2Web(url);
        }



        private void combox_sd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_sd);
        }

        private void combox_dk_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_dk);
        }

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_dk.ComboxText(), combox_field);
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.SaveDialogExcel();
        }
    }
}
