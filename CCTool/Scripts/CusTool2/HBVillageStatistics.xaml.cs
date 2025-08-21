using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Vml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Brushes = System.Windows.Media.Brushes;

namespace CCTool.Scripts.CusTool2
{
    /// <summary>
    /// Interaction logic for HBVillageStatistics.xaml
    /// </summary>
    public partial class HBVillageStatistics : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public HBVillageStatistics()
        {
            InitializeComponent();

            combox_area.Items.Add("投影面积");
            combox_area.Items.Add("椭球面积");
            combox_area.SelectedIndex = 0;

            textExcelPath.Text = $@"{Project.Current.HomeFolderPath}\导出湖北村规指标表";
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "湖北省村规结构调整表(D)";

        // 点击打开按钮，选择输出的Excel文件位置
        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            // 打开Excel文件
            string path = UITool.OpenDialogFolder();
            // 将Excel文件的路径置入【textExcelPath】
            textExcelPath.Text = path;
        }

        // 添加要素图层的所有字符串字段到combox中
        private void combox_field_DropDown(object sender, EventArgs e)
        {
            // 将图层字段加入到Combox列表中
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_bmField);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);

        }

        // 运行
        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                string init_gdb = Project.Current.DefaultGeodatabasePath;
                string init_foder = Project.Current.HomeFolderPath;

                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string field_bm = combox_bmField.ComboxText();

                string fc_path_gh = combox_fc_gh.ComboxText();
                string field_bm_gh = combox_bmField_gh.ComboxText();

                string excel_folder = textExcelPath.Text;

                string czcField = "CZCSXM";   //  城镇村属性码

                string area_type = combox_area.Text[..2];   // 面积类型

                // 判断参数是否选择完全
                if (fc_path == "" || field_bm == "" || fc_path_gh == "" || field_bm_gh == "" || excel_folder == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                await QueuedTask.Run(() =>
                {
                    if (!GisTool.IsHaveFieldInTarget(fc_path, czcField))
                    {
                        MessageBox.Show("现状图层缺少【CZCSXM】字段！");
                    }
                    if (!GisTool.IsHaveFieldInTarget(fc_path_gh, czcField))
                    {
                        MessageBox.Show("规划图层缺少【CZCSXM】字段！");
                    }
                });

                // 创建目标文件夹
                if (!Directory.Exists(excel_folder))
                {
                    Directory.CreateDirectory(excel_folder);
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    string excelName = "村域国土空间用途结构调整表";
                    string excelMapper = "用地用海代码_村庄功能";

                    pw.AddMessageStart("复制模板");
                    // 复制嵌入资源中的Excel文件
                    string excel_path = $@"{excel_folder}\村域国土空间用途结构调整表.xlsx";
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.湖北村规.{excelName}.xlsx", excel_path);
                    string excel_sheet = excel_path + @"\Sheet1$";
                    // 复制映射表
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.湖北村规.{excelMapper}.xlsx", @$"{init_foder}\{excelMapper}.xlsx");

                    pw.AddMessageMiddle(10, "村域_现状用地汇总");
                    /// 全域现状用地
                    // 复制要素
                    string xzyd = $@"{init_gdb}\xzyd";
                    Arcpy.CopyFeatures(fc_path, xzyd);
                    // 计算面积
                    string mjField = "计算面积字段";
                    GisTool.AddField(xzyd, mjField, FieldType.Double);
                    string block = area_type == "投影" ? "!shape.area!" : "!shape.geodesicarea!";
                    Arcpy.CalculateField(xzyd, mjField, block);
                    // 初步汇总
                    // 调用GP工具【汇总】
                    string output_table = init_gdb + @"\output_table";
                    Arcpy.Statistics(xzyd, output_table, $"{mjField} SUM", $"{czcField};{field_bm}");
                    // 添加村庄功能字段
                    string gnField = "功能分类";
                    GisTool.AddField(output_table, gnField, FieldType.String);

                    // 村庄功能映射
                    string mapper = @$"{init_foder}\{excelMapper}.xlsx\sheet1$";
                    ComboTool.AttributeMapper(output_table, field_bm, gnField, mapper);
                    // 区分202、203等属性
                    Segment(output_table, gnField, "现状");

                    // 汇总指标
                    Dictionary<string, double> dict = ComboTool.StatisticsPlus(output_table, gnField, $"SUM_{mjField}", "合计", 10000);
                    // 获取比例
                    Dictionary<string, double> dict_per = GetPesent(dict);

                    // 属性映射现状用地
                    ExcelTool.AttributeMapperDouble(excel_sheet, 0, 1, dict, 4);
                    ExcelTool.AttributeMapperDouble(excel_sheet, 0, 2, dict_per, 4);

                    pw.AddMessageMiddle(20, "村域_规划用地汇总");
                    /// 全域规划用地
                    // 复制要素
                    string ghyd = $@"{init_gdb}\ghyd";
                    Arcpy.CopyFeatures(fc_path_gh, ghyd);
                    // 计算面积
                    GisTool.AddField(ghyd, mjField, FieldType.Double);
                    Arcpy.CalculateField(ghyd, mjField, block);
                    // 初步汇总
                    // 调用GP工具【汇总】
                    string output_table_gh = init_gdb + @"\output_table_gh";
                    Arcpy.Statistics(ghyd, output_table_gh, $"{mjField} SUM", $"{czcField};{field_bm_gh}");
                    // 添加村庄功能字段
                    GisTool.AddField(output_table_gh, gnField, FieldType.String);

                    // 村庄功能映射
                    ComboTool.AttributeMapper(output_table_gh, field_bm_gh, gnField, mapper);
                    // 区分202、203等属性
                    Segment(output_table_gh, gnField, "规划");

                    // 汇总指标
                    dict = ComboTool.StatisticsPlus(output_table_gh, gnField, $"SUM_{mjField}", "合计", 10000);
                    // 获取比例
                    dict_per = GetPesent(dict);

                    // 属性映射现状用地
                    ExcelTool.AttributeMapperDouble(excel_sheet, 0, 3, dict, 4);
                    ExcelTool.AttributeMapperDouble(excel_sheet, 0, 4, dict_per, 4);

                    // 删除空行
                    ExcelTool.DeleteNullRow(excel_sheet, new List<int>() { 1, 3 }, 4);


                    /// 村庄建设边界内_现状用地
                    string excelName2 = "村庄建设边界内用地结构规划表";
                    string excelMapper2 = "用地用海代码_村庄功能_边界内";

                    // 复制嵌入资源中的Excel文件
                    string excel_path2 = $@"{excel_folder}\村庄建设边界内用地结构规划表.xlsx";
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.湖北村规.{excelName2}.xlsx", excel_path2);
                    string excel_sheet2 = excel_path2 + @"\Sheet1$";
                    // 复制映射表
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.湖北村规.{excelMapper2}.xlsx", @$"{init_foder}\{excelMapper2}.xlsx");


                    pw.AddMessageMiddle(20, "村庄建设边界内_现状用地汇总");
                    // 选择边界内的
                    string output_table_bjn = @$"{init_gdb}\output_table_bjn";
                    Arcpy.TableSelect(output_table, output_table_bjn, "功能分类 = '村庄用地' Or CZCSXM LIKE '%203%'");

                    string mapper2 = @$"{init_foder}\{excelMapper2}.xlsx\sheet1$";
                    // 村庄功能映射
                    ComboTool.AttributeMapper(output_table_bjn, field_bm, gnField, mapper2);

                    // 汇总指标
                    dict = ComboTool.StatisticsPlus(output_table_bjn, gnField, $"SUM_{mjField}", "合计", 10000);
                    // 获取比例
                    dict_per = GetPesent(dict);

                    // 属性映射现状用地
                    ExcelTool.AttributeMapperDouble(excel_sheet2, 10, 2, dict, 4);
                    ExcelTool.AttributeMapperDouble(excel_sheet2, 10, 3, dict_per, 4);


                    /// 村庄建设边界内_规划用地
                    pw.AddMessageMiddle(20, "村庄建设边界内_规划用地汇总");
                    // 选择边界内的
                    string output_table_gh_bjn = @$"{init_gdb}\output_table_gh_bjn";
                    Arcpy.TableSelect(output_table_gh, output_table_gh_bjn, "功能分类 = '村庄用地'");
                    // 村庄功能映射
                    ComboTool.AttributeMapper(output_table_gh_bjn, field_bm_gh, gnField, mapper2);

                    // 汇总指标
                    dict = ComboTool.StatisticsPlus(output_table_gh_bjn, gnField, $"SUM_{mjField}", "合计", 10000);
                    // 获取比例
                    dict_per = GetPesent(dict);

                    // 属性映射现状用地
                    ExcelTool.AttributeMapperDouble(excel_sheet2, 10, 4, dict, 4);
                    ExcelTool.AttributeMapperDouble(excel_sheet2, 10, 5, dict_per, 4);

                    // 删除空行
                    ExcelTool.DeleteNullRow(excel_sheet2, new List<int>() { 2, 4 }, 4);
                    // 删除列
                    ExcelTool.DeleteCol(excel_sheet2, 10);


                    // 删除中间数据
                    File.Delete($@"{init_foder}\用地用海代码_村庄功能.xlsx");
                    File.Delete($@"{init_foder}\用地用海代码_村庄功能_边界内.xlsx");

                    List<string> fcs = new List<string>() { "output_table", "output_table_bjn", "output_table_gh", "output_table_gh_bjn", "xzyd", "ghyd" };

                    foreach (string fc in fcs)
                    {
                        Arcpy.Delect($@"{init_gdb}\{fc}");
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

        // 区分202、203等属性
        public void Segment(string tablePath, string gnField, string zt)
        {
            Table table = tablePath.TargetTable();
            using RowCursor rowCursor = table.Search();
            while (rowCursor.MoveNext())
            {
                using Row row = rowCursor.Current;
                // 获取value
                string czcsxm = row["CZCSXM"]?.ToString() ?? "";
                // 判断, 如果是城镇
                if (czcsxm == "201" || czcsxm == "202" || czcsxm == "201A" || czcsxm == "202A" || czcsxm == "201a" || czcsxm == "202a")
                {
                    row[gnField] = "城镇用地";
                }

                if (zt == "现状")
                {
                    // 判断, 如果是村庄
                    if (czcsxm == "203" || czcsxm == "203A" || czcsxm == "203a")
                    {
                        row[gnField] = "村庄用地";
                    }
                }
                row.Store();
            }
        }



        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/144241347";
            UITool.Link2Web(url);
        }

        private void combox_fc_gh_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc_gh);
        }

        private void combox_bmField_gh_DropDown(object sender, EventArgs e)
        {
            // 将图层字段加入到Combox列表中
            UITool.AddTextFieldsToComboxPlus(combox_fc_gh.ComboxText(), combox_bmField_gh);
        }

        // 比例
        private Dictionary<string, double> GetPesent(Dictionary<string, double> dict)
        {
            Dictionary<string, double> result = new Dictionary<string, double>();
            double total = dict["合计"];

            foreach (var item in dict)
            {
                string key = item.Key;
                double value = item.Value;

                result.Add(key, value / total * 100);
            }

            return result;
        }

    }
}
