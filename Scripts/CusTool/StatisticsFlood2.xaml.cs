using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
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

namespace CCTool.Scripts.CusTool
{
    /// <summary>
    /// Interaction logic for StatisticsFlood2.xaml
    /// </summary>
    public partial class StatisticsFlood2 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public StatisticsFlood2()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "洪水四线分析加强版";

        private void combox_sd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_sd);
            textExcelPath.Text = Project.Current.HomeFolderPath + @"\洪水四线分析.xlsx";
        }

        private void combox_df_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_df);
        }

        private void combox_lsx_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_lsx);
        }

        private void combox_xhx_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_xhx);
        }

        private void combox_tcx_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_tcx);
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

                string lsx = combox_lsx.ComboxText();
                string xhx = combox_xhx.ComboxText();
                string tcx = combox_tcx.ComboxText();

                // 输出路径
                string excel_path = textExcelPath.Text;

                // 判断参数是否选择完全
                if (sd == "" || df == "" || excel_path == "")
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
                    List<string> lines = new List<string>() { sd, df, lsx, xhx, tcx };
                    // 检查数据
                    List<string> errs = CheckData(lines, sd);
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
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.三上数据地类统计表.xlsx", excel_path);
                    
                    pw.AddMessageMiddle(10, "三调地类统计");

                    // 复制三调
                    string sd_copy = $@"{def_gdb}\sd_copy";
                    string gnField = "归纳";
                    Arcpy.CopyFeatures(sd, sd_copy);
                    Arcpy.DeleteField(sd_copy, "DLBM", "KEEP_FIELDS");
                    Arcpy.AddField(sd_copy, gnField, "TEXT");
                    // 归纳
                    ComboTool.AttributeMapper(sd_copy, "DLBM", gnField, excel_mapper + @"\sheet1$");

                    pw.AddMessageMiddle(10, "【生成总表】", Brushes.Blue);
                    // 汇总并分割
                    string new_gdb = StatisticsAndClip(sd_copy, df, xField, tField, gnField);
                    // 收集分割后的地块
                    List<string> list_area_detail = new_gdb.GetFeatureClassAndTablePath();
                    // 逐个分析
                    int start_row = 4;
                    string new_sheet_path = $@"{excel_path}\山东黄河两岸堤防内三上数据地类分类统计表$";
                    foreach (var area_detail in list_area_detail)
                    {
                        if ((start_row - 3) % 10 == 0)
                        {
                            pw.AddMessageMiddle(0, $"第{start_row - 3}行……", Brushes.Gray);
                        }
                        // 获取县名和滩区名
                        string xName = area_detail.TargetCellValue(xField);
                        string tName = area_detail.TargetCellValue(tField);

                        // 复制行
                        ExcelTool.CopyRows(new_sheet_path, 3, start_row);

                        // 写入标记
                        string index = (start_row - 3).ToString();
                        ExcelTool.WriteCell(new_sheet_path, start_row, 1, index);
                        ExcelTool.WriteCell(new_sheet_path, start_row, 2, xName);
                        if (tName == "最后")   // 合计行
                        {
                            ExcelTool.WriteCell(new_sheet_path, start_row, 3, "合计");
                        }
                        else
                        {
                            ExcelTool.WriteCell(new_sheet_path, start_row, 3, tName);
                        }

                        // 获取dict
                        Dictionary<string, double> dict = GisTool.GetDictFromPathDouble(area_detail, gnField, "MJJ");
                        // 属性映射到Excel
                        ExcelTool.AttributeMapperColDouble(new_sheet_path, 2, start_row, dict);

                        start_row++;
                    }
                    // 整理
                    ExcelTool.DeleteRow(new_sheet_path, 3);
                    ExcelTool.MergeSameCol(new_sheet_path, 2, 3);


                    // 汇总三线
                    if (lsx != "")
                    {
                        pw.AddMessageMiddle(10, "【临水线表】", Brushes.Blue);
                        string sheet_name = "山东黄河两岸堤防内三上数据与临水线、滩区叠加地类分类统计表";
                        ClipStatistics(sd_copy, lsx, df, gnField, excel_path, sheet_name, xField, tField, pw);
                    }

                    if (xhx != "")
                    {
                        pw.AddMessageMiddle(20, "【秋汛线表】", Brushes.Blue);
                        string sheet_name = "山东黄河两岸堤防内三上数据与秋汛线、滩区叠加地类分类统计表";
                        ClipStatistics(sd_copy, xhx, df, gnField, excel_path, sheet_name, xField, tField, pw);
                    }

                    if (tcx != "")
                    {
                        pw.AddMessageMiddle(20, "【滩槽分界线表】", Brushes.Blue);
                        string sheet_name = "山东黄河两岸堤防内三上数据与滩槽分界线、滩区叠加地类分类统计表";
                        ClipStatistics(sd_copy, tcx, df, gnField, excel_path, sheet_name, xField, tField, pw);
                    }


                    // 删除中间数据
                    Arcpy.Delect(sd_copy);
                    Arcpy.Delect($@"{def_gdb}\sd_identity");
                    Arcpy.Delect($@"{def_gdb}\outTable");
                    Arcpy.Delect($@"{def_gdb}\outTable2");
                    Arcpy.Delect($@"{def_gdb}\outTable3");
                    Arcpy.Delect($@"{def_gdb}\outTable4");

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

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "";
            UITool.Link2Web(url);
        }

        // 分区裁剪汇总
        public void ClipStatistics(string sd_copy, string zone, string df, string gnField, string excel_path, string sheet_name, string xField, string tField, ProcessWindow pw)
        {
            string def_gdb = Project.Current.DefaultGeodatabasePath;

            // 添加标记字段
            string bjField = "标记";
            Arcpy.AddField(zone, bjField, "TEXT");
            Arcpy.CalculateField(zone, bjField, "'线内'");
            // 标识
            string sd_identity = $@"{def_gdb}\sd_identity";
            Arcpy.Identity(sd_copy, zone, sd_identity);
            Arcpy.CalculateField(sd_identity, gnField, $"ss(!{gnField}! , !{bjField}!)", "def ss(a,b):\r\n    if b == \"\":\r\n        return a +\"线外\"\r\n    else:\r\n        return a + b");

            // 汇总并分割
            string new_gdb = StatisticsAndClip2(sd_identity, df, xField, tField, gnField);

            // 收集分割后的地块
            List<string> list_area_detail = new_gdb.GetFeatureClassAndTablePath();
            // 逐个分析
            int start_row = 6;
            string new_sheet_path = $@"{excel_path}\{sheet_name}$";
            foreach (var area_detail in list_area_detail)
            {
                if ((start_row - 5) % 10 == 0)
                {
                    pw.AddMessageMiddle(0, $"第{start_row - 5}行……", Brushes.Gray);
                }
                // 获取县名和滩区名
                string xName = area_detail.TargetCellValue(xField);
                string tName = area_detail.TargetCellValue(tField);

                // 复制行
                ExcelTool.CopyRows(new_sheet_path, 5, start_row);

                // 写入标记
                string index = (start_row - 5).ToString();
                ExcelTool.WriteCell(new_sheet_path, start_row, 0, index);
                ExcelTool.WriteCell(new_sheet_path, start_row, 1, xName);
                if (tName == "最后")   // 合计行
                {
                    ExcelTool.WriteCell(new_sheet_path, start_row, 2, "合计");
                }
                else
                {
                    ExcelTool.WriteCell(new_sheet_path, start_row, 2, tName);
                }

                // 获取dict
                Dictionary<string, double> dict = GisTool.GetDictFromPathDouble(area_detail, gnField, "MJJ");
                // 属性映射到Excel
                ExcelTool.AttributeMapperColDouble(new_sheet_path, 1, start_row, dict);

                start_row++;
            }
            // 整理
            ExcelTool.DeleteRow(new_sheet_path, new List<int>() { 5, 1 } );
            ExcelTool.MergeSameCol(new_sheet_path, 1, 4);

            Arcpy.DeleteField(zone, bjField);
        }

        // 汇总并分割
        public string StatisticsAndClip(string sd_copy, string df, string xField, string tField, string gnField)
        {
            string def_gdb = Project.Current.DefaultGeodatabasePath;
            string def_path = Project.Current.HomeFolderPath;
            //交集制表
            string outTable = $@"{def_gdb}\outTable";
            string outTable2 = $@"{def_gdb}\outTable2";
            Arcpy.TabulateIntersection(df, new List<string>() { xField, tField }, sd_copy, outTable, gnField);
            Arcpy.TabulateIntersection(df, new List<string>() { xField }, sd_copy, outTable2, gnField);
            // 追加
            Arcpy.Append(outTable2, outTable);
            Arcpy.CalculateField(outTable, tField, $"ss(!{tField}!)", "def ss(a):\r\n    if a is None:\r\n        return \"最后\"\r\n    else:\r\n        return a");
            // 面积单位换算
            string mjField = "MJJ";
            Arcpy.AddField(outTable, mjField, "DOUBLE");
            Arcpy.CalculateField(outTable, mjField, "!shape_area!/10000*15/10000");

            // 放到新建的数据库中
            string filePath = $@"{def_path}\CC配置文件(勿删)";
            if (!Directory.Exists(filePath)) { Directory.CreateDirectory(filePath); }
            string gdbName = "临时数据库";
            Arcpy.CreateFileGDB(filePath, gdbName);
            string new_gdb = $@"{filePath}\{gdbName}.gdb";
            // 清理一下
            GisTool.ClearGDBItem(new_gdb);
            // 拆分
            Arcpy.SplitByAttributes(outTable, new_gdb, $"{xField};{tField}");

            return new_gdb;
        }

        // 汇总并分割_分线内外
        public string StatisticsAndClip2(string sd_copy, string df, string xField, string tField, string gnField)
        {
            string def_gdb = Project.Current.DefaultGeodatabasePath;
            string def_path = Project.Current.HomeFolderPath;
            //交集制表
            string outTable = $@"{def_gdb}\outTable";
            string outTable3 = $@"{def_gdb}\outTable3";
            string outTable4 = $@"{def_gdb}\outTable4";
            Arcpy.TabulateIntersection(df, new List<string>() { xField, tField }, sd_copy, outTable3, gnField);
            Arcpy.TabulateIntersection(df, new List<string>() { xField }, sd_copy, outTable4, gnField);
            // 追加
            Arcpy.CalculateField(outTable, gnField, $"ss(!{gnField}!)", "def ss(a):\r\n    if a[-2:] == \"合计\":\r\n        return a\r\n    else:\r\n        return a + \"合计\"");
            Arcpy.Append(outTable4, outTable3);
            Arcpy.Append(outTable, outTable3);
            Arcpy.CalculateField(outTable3, tField, $"ss(!{tField}!)", "def ss(a):\r\n    if a is None:\r\n        return \"最后\"\r\n    else:\r\n        return a");
            // 面积单位换算
            string mjField = "MJJ";
            Arcpy.AddField(outTable3, mjField, "DOUBLE");
            Arcpy.CalculateField(outTable3, mjField, "!shape_area!/10000*15/10000");

            // 放到新建的数据库中
            string filePath = $@"{def_path}\CC配置文件(勿删)";
            if (!Directory.Exists(filePath)) { Directory.CreateDirectory(filePath); }
            string gdbName = "临时数据库";
            Arcpy.CreateFileGDB(filePath, gdbName);
            string new_gdb = $@"{filePath}\{gdbName}.gdb";
            // 清理一下
            GisTool.ClearGDBItem(new_gdb);
            // 拆分
            Arcpy.SplitByAttributes(outTable3, new_gdb, $"{xField};{tField}");

            return new_gdb;
        }


        private List<string> CheckData(List<string> lines, string sd)
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
    }
}
