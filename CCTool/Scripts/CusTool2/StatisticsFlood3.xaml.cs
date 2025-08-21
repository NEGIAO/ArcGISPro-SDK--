using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
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
    /// Interaction logic for StatisticsFlood3.xaml
    /// </summary>
    public partial class StatisticsFlood3 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public StatisticsFlood3()
        {
            InitializeComponent();

            //UITool.InitFieldToComboxPlus(combox_xName, "县名", "string");
            //UITool.InitFieldToComboxPlus(combox_tName, "滩区名", "string");
            //UITool.InitFieldToComboxPlus(combox_tType, "类别", "string");

            //UITool.InitFieldToComboxPlus(combox_sd, "现状用地", "string");
            //UITool.InitFieldToComboxPlus(combox_df, "滩区面", "string");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "洪水分析(老嫩滩)";

        private void combox_sd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_sd);
            textExcelPath.Text = Project.Current.HomeFolderPath + @"\洪水分析(老嫩滩).xlsx";
        }

        private void combox_df_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_df);
        }


        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.SaveDialogExcel();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                string def_gdb = Project.Current.DefaultGeodatabasePath;
                string def_path = Project.Current.HomeFolderPath;
                // 获取分析要素
                string sd = combox_sd.ComboxText();
                string df = combox_df.ComboxText();

                string xField = combox_xName.ComboxText();
                string tField = combox_tName.ComboxText();
                string tType = combox_tType.ComboxText();

                // 输出路径
                string excel_path = textExcelPath.Text;

                // 判断参数是否选择完全
                if (sd == "" || df == "" || xField == "" || tField == "" || tType == "")
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
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(sd);
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
                    string excel_mapper = $@"{def_path}\三调用地自转换_洪水2.xlsx";
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.三调用地自转换_洪水2.xlsx", excel_mapper);
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.三上数据地类统计表20240904模板(老嫩滩).xlsx", excel_path);

                    pw.AddMessageMiddle(10, "三调地类统计");

                    // 交集取反
                    List<string> fields = new List<string>() { xField, tField, tType };
                    string outTable = $@"{def_gdb}\outTable";
                    Arcpy.TabulateIntersection(df, fields, sd, outTable, "DLBM");
                    // 换单位_万亩
                    Arcpy.CalculateField(outTable, "AREA", "!AREA!/10000*15/10000");
                    // 归纳用地类别
                    string gnField = "归纳";
                    Arcpy.AddField(outTable, gnField, "TEXT");
                    ComboTool.AttributeMapper(outTable, "DLBM", gnField, excel_mapper + @"\sheet1$");

                    string outTable2 = $@"{def_gdb}\outTable2";
                    Arcpy.Statistics(outTable, outTable2, $"AREA SUM", $"{xField};{tField};{tType};{gnField}");

                    pw.AddMessageMiddle(20, "获取指标");
                    List<string> fdList = new List<string>() { xField, tField, tType, gnField };
                    var areaValues = GisTool.GetDictListFromPathDouble(outTable2, fdList, "SUM_AREA");


                    pw.AddMessageMiddle(10, "写入统计表");
                    // 获取工作薄、工作表
                    string excelFile = ExcelTool.GetPath(excel_path);
                    int sheetIndex = ExcelTool.GetSheetIndex(excel_path);
                    // 打开工作薄
                    Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                    // 打开工作表
                    Worksheet sheet = wb.Worksheets[sheetIndex];
                    // 逐行处理
                    for (int i = 5; i <= sheet.Cells.MaxDataRow; i++)
                    {
                        for (int j = 3; j < sheet.Cells.MaxDataColumn; j++)
                        {
                            //  获取目标cell
                            Cell inCell = sheet.Cells[i, j];
                            //  获取对照cell
                            Cell xNameCell = sheet.Cells[i, 34];     // 县名
                            Cell tNameCell = sheet.Cells[i, 33];     // 滩区名
                            Cell tTypeCell = sheet.Cells[4, j];         // 老嫩滩
                            Cell ydCell = sheet.Cells[1, j];               // 用地名称
                            // 属性映射
                            if (ydCell is not null && tNameCell is not null && tTypeCell is not null && xNameCell is not null)
                            {
                                foreach (var values in areaValues)
                                {
                                    // 其余的情况
                                    if (tNameCell.StringValue == "其余")
                                    {
                                        if (tNameCell.StringValue == values.Key[1] && tTypeCell.StringValue == values.Key[2] && ydCell.StringValue == values.Key[3])
                                        {
                                            inCell.Value = values.Value;   // 赋值
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        if (xNameCell.StringValue == values.Key[0] && tNameCell.StringValue == values.Key[1] && tTypeCell.StringValue == values.Key[2] && ydCell.StringValue == values.Key[3])
                                        {
                                            inCell.Value = values.Value;   // 赋值
                                            break;
                                        }
                                    }
                                    
                                }
                            }
                        }
                    }
                    // 保存
                    wb.Save(excelFile);
                    wb.Dispose();

                    // 删除多余行列
                    ExcelTool.DeleteRow(excel_path, 1);
                    ExcelTool.DeleteCol(excel_path, new List<int>() { 34,33});

                    // 删除中间数据
                    Arcpy.Delect(outTable);
                    Arcpy.Delect(outTable2);

                    File.Delete(excel_mapper);

                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }




        private List<string> CheckData(string sd)
        {
            List<string> result = new List<string>();


            // 检查DLBM的字段值
            string result_value = CheckTool.CheckFieldValue(sd, "DLBM", GlobalData.dic_sdAll.Keys.ToList());
            if (result_value != "")
            {
                result.Add(result_value);
            }

            // 检查是否正常提取Excel
            string result_excel = CheckTool.CheckExcelPick();
            if (result_excel != "")
            {
                result.Add(result_excel);
            }

            return result;
        }

        private void combox_xName_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_df.ComboxText(), combox_xName);
        }

        private void combox_tName_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_df.ComboxText(), combox_tName);
        }

        private void combox_tType_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_df.ComboxText(), combox_tType);
        }
    }
}
