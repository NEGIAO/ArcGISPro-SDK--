using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Framework.Utilities;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using MathNet.Numerics;
using MathNet.Numerics.LinearAlgebra.Factorization;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;
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
using Range = Aspose.Cells.Range;
using Row = ArcGIS.Core.Data.Row;
using Table = ArcGIS.Core.Data.Table;

namespace CCTool.Scripts.CusTool
{
    /// <summary>
    /// Interaction logic for SXVillageTable.xaml
    /// </summary>
    public partial class SXVillageTable : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public SXVillageTable()
        {
            InitializeComponent();

            // 初始化combox
            //UITool.InitFeatureLayerToComboxPlus(combox_fc, "村子数据1村");
            //textFolderPath.Text = @"C:\Users\Administrator\Desktop\新输出";


            UITool.InitFieldToComboxPlus(combox_nameField, "CZMC", "string");
            UITool.InitFieldToComboxPlus(combox_bmField_xz, "XZBM", "string");
            UITool.InitFieldToComboxPlus(combox_bmField_gh, "GHBM", "string");
            UITool.InitFieldToComboxPlus(combox_czcField_xz, "CZCSXM", "string");
            UITool.InitFieldToComboxPlus(combox_czcField_gh, "GHCZC", "string");

            // 初始化combox
            combox_unit.Items.Add("平方米");
            combox_unit.Items.Add("公顷");
            combox_unit.Items.Add("平方公里");
            combox_unit.Items.Add("亩");
            combox_unit.SelectedIndex = 1;

            combox_area.Items.Add("投影面积");
            combox_area.Items.Add("椭球面积");
            combox_area.SelectedIndex = 1;

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "山西省村规结构调整表(亦求长生亦求你)";


        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void combox_nameField_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_nameField);
        }

        private void combox_bmField_gh_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_bmField_gh);
        }

        private void combox_bmField_xz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_bmField_xz);
        }

        private void combox_czcField_xz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_czcField_xz);
        }

        private void combox_czcField_gh_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_czcField_gh);
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textFolderPath.Text = UITool.OpenDialogFolder();
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/139373689";
            UITool.Link2Web(url);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                string init_gdb = Project.Current.DefaultGeodatabasePath;
                string init_foder = Project.Current.HomeFolderPath;

                // 参数获取
                string in_fc = combox_fc.ComboxText();
                string excel_folder = textFolderPath.Text;

                string nameField = combox_nameField.ComboxText();
                string bmField_xz = combox_bmField_xz.ComboxText();
                string bmField_gh = combox_bmField_gh.ComboxText();
                string czcField_xz = combox_czcField_xz.ComboxText();
                string czcField_gh = combox_czcField_gh.ComboxText();

                string area_type = combox_area.Text[..2];
                string unit = combox_unit.Text;

                // 判断参数是否选择完全
                if (in_fc == "" || excel_folder == "" || nameField == "" || bmField_xz == "" || bmField_gh == ""
                    || czcField_xz == "" || czcField_gh == "")
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
                    // 汇总用地
                    string excelMapper = "【山西】用地用海代码_村庄功能.xlsx";

                    // 复制映射表
                    string targetMapper = @$"{excel_folder}\{excelMapper}";
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.{excelMapper}", targetMapper);

                    pw.AddMessageStart("计算面积");

                    // 单位系数设置
                    double unit_xs = unit switch
                    {
                        "平方米" => 1,
                        "公顷" => 10000,
                        "平方公里" => 1000000,
                        "亩" => 666.66667,
                        _ => 1,
                    };
                    // 计算面积
                    string mjField = "jsmj";
                    Arcpy.AddField(in_fc, mjField, "DOUBLE");
                    if (area_type == "投影")
                    {
                        Arcpy.CalculateField(in_fc, mjField, $"!shape.area!");
                    }
                    else
                    {
                        Arcpy.CalculateField(in_fc, mjField, $"!shape.geodesicarea!");
                    }
                    pw.AddMessageMiddle(0, "初步汇总");
                    // 调用GP工具【汇总】
                    string staTable = $@"{init_gdb}\SX_table";
                    Arcpy.Statistics(in_fc, staTable, $"{mjField} SUM", $"{nameField};{bmField_xz};{bmField_gh};{czcField_xz};{czcField_gh}");

                    pw.AddMessageMiddle(10, "村庄功能映射");
                    // 添加村庄功能字段
                    string gn_xz = "gn_xz";
                    string gn_gh = "gn_gh";
                    string gn_xz2 = "gn_xz2";
                    string gn_gh2 = "gn_gh2";
                    string gn_xz3 = "gn_xz3";
                    string gn_gh3 = "gn_gh3";
                    string gn_xz4 = "gn_xz4";
                    string gn_gh4 = "gn_gh4";
                    Arcpy.AddField(staTable, gn_xz, "TEXT");
                    Arcpy.AddField(staTable, gn_gh, "TEXT");
                    Arcpy.AddField(staTable, gn_xz2, "TEXT");
                    Arcpy.AddField(staTable, gn_gh2, "TEXT");
                    Arcpy.AddField(staTable, gn_xz3, "TEXT");
                    Arcpy.AddField(staTable, gn_gh3, "TEXT");
                    Arcpy.AddField(staTable, gn_xz4, "TEXT");
                    Arcpy.AddField(staTable, gn_gh4, "TEXT");
                    // 村庄功能映射
                    string mapper = @$"{targetMapper}\村域$";
                    string mapper2 = @$"{targetMapper}\村域2$";
                    string mapper3 = @$"{targetMapper}\村域3$";
                    string mapper4 = @$"{targetMapper}\村庄$";
                    ComboTool.AttributeMapper(staTable, bmField_xz, gn_xz, mapper);
                    ComboTool.AttributeMapper(staTable, bmField_gh, gn_gh, mapper);
                    ComboTool.AttributeMapper(staTable, bmField_xz, gn_xz2, mapper2);
                    ComboTool.AttributeMapper(staTable, bmField_gh, gn_gh2, mapper2);
                    ComboTool.AttributeMapper(staTable, czcField_xz, gn_xz3, mapper3);
                    ComboTool.AttributeMapper(staTable, czcField_gh, gn_gh3, mapper3);
                    ComboTool.AttributeMapper(staTable, bmField_xz, gn_xz4, mapper4);
                    ComboTool.AttributeMapper(staTable, bmField_gh, gn_gh4, mapper4);

                    // 把城镇村指标合并到功能字段
                    string code1 = "def ss(a,b):\r\n    if b is None:\r\n        return a\r\n    else:\r\n        return b";
                    string code2 = "def ss(a,b):\r\n    if b is None:\r\n        return a\r\n    else:\r\n        return b";
                    string code3 = "def ss(a,b,c,d):\r\n    va1 = a\r\n    va2 = b\r\n    if va1 is None:\r\n        va1=c\r\n    if va2 is None:\r\n        va2=d\r\n    return va1+\"+\"+va2";
                    string code4 = "def ss(a,b):\r\n    if a in(\"203\", \"203A\"):\r\n        return b\r\n    else:\r\n        return \"\"";

                    string code5 = "def ss(a,b):\r\n    if b in (\"201\",\"201A\",\"202\",\"202A\",\"203\",\"203A\"):\r\n        return None\r\n    else:\r\n        return a";

                    pw.AddMessageMiddle(10, $"计算字段，方便后续统计");
                    Arcpy.CalculateField(staTable, gn_xz2, $"ss(!{gn_xz2}!,!{czcField_xz}!)", code5);
                    Arcpy.CalculateField(staTable, gn_gh2, $"ss(!{gn_gh2}!,!{czcField_gh}!)", code5);

                    Arcpy.CalculateField(staTable, gn_xz, $"ss(!{gn_xz}!,!{gn_xz3}!)", code1);
                    Arcpy.CalculateField(staTable, gn_gh, $"ss(!{gn_gh}!,!{gn_gh3}!)", code2);
                    Arcpy.CalculateField(staTable, gn_xz3, $"ss(!{gn_xz}!,!{gn_gh}!,!{gn_xz2}!,!{gn_gh2}!)", code3);
                    Arcpy.CalculateField(staTable, gn_xz4, $"ss(!{czcField_xz}!,!{gn_xz4}!)", code4);
                    Arcpy.CalculateField(staTable, gn_gh4, $"ss(!{czcField_gh}!,!{gn_gh4}!)", code4);
                    // 获取所有村庄名称
                    List<string> vgList = staTable.GetFieldValues(nameField);
                    // 分村汇总
                    foreach (var vg in vgList)
                    {
                        pw.AddMessageMiddle(10, $"{vg}", Brushes.Green);
                        pw.AddMessageMiddle(0, $"生成村域表", Brushes.Gray);
                        string tb = $@"{init_gdb}\{vg}_table";
                        Arcpy.TableSelect(staTable, tb, $"{nameField} = '{vg}'");
                        // 复制嵌入资源中的Excel文件
                        string excelName = "【山西】国土空间用途结构调整表.xlsx";
                        string targetExcel = @$"{excel_folder}\{vg}.xlsx";
                        DirTool.CopyResourceFile(@$"CCTool.Data.Excel.{excelName}", targetExcel);

                        // 汇总指标
                        List<string> bmList_xz = new List<string>() { gn_xz, gn_xz2 };
                        List<string> bmList_gh = new List<string>() { gn_gh, gn_gh2 };
                        string total = "国土面积";
                        Dictionary<string, double> dic_xz = ComboTool.MultiStatisticsToDic(tb, $"Sum_{mjField}", bmList_xz, total, unit_xs);
                        Dictionary<string, double> dic_gh = ComboTool.MultiStatisticsToDic(tb, $"Sum_{mjField}", bmList_gh, total, unit_xs);
                        // 统计用地变化量
                        // 初始化用地集合
                        List<string> ydList = new List<string>()
                        {
                            "耕地","园地","林地","草地","湿地","农村道路","设施农用地",
                            "种植设施建设用地","畜禽养殖设施建设用地","城镇","村庄","区域基础设施用地","一类工业用地","二类工业用地",
                            "三类工业用地","采矿用地","一类物流仓储用地","二类物流仓储用地","三类物流仓储用地",
                            "储备库用地","其他用地","其他土地","陆地水域",
                        };
                        Dictionary<string, double> dic_increase = new Dictionary<string, double>();   // 增加量
                        Dictionary<string, double> dic_reduce = new Dictionary<string, double>();     // 减少量
                        foreach (var yd in ydList)
                        {
                            dic_increase.Add(yd, 0);
                            dic_reduce.Add(yd, 0);
                        }
                        // 遍历表格，更新增减量指标
                        Table table = tb.TargetTable();
                        // 逐行找出错误
                        using RowCursor rowCursor = table.Search();
                        while (rowCursor.MoveNext())
                        {
                            using Row row = rowCursor.Current;
                            // 获取value
                            double mj = double.Parse(row["SUM_jsmj"].ToString()) / unit_xs;
                            var va = row[gn_xz3];
                            if (va != null)
                            {
                                // 获取现状规划名称
                                string xz = va.ToString().Split('+')[0];
                                string gh = va.ToString().Split('+')[1];
                                // 分析增减
                                if (xz != gh)
                                {
                                    dic_reduce[xz] += mj;    // 减少量
                                    dic_increase[gh] += mj;    // 增加量
                                    // 如果是种植或养殖用地
                                    List<string> nyd = new List<string>() { "种植设施建设用地", "养殖设施建设用地" };
                                    bool isXZ = nyd.Contains(xz);
                                    bool isGH = nyd.Contains(gh);

                                    if ((isXZ && !isGH))
                                    {
                                        xz = "设施农用地";
                                        dic_reduce[xz] += mj;    // 减少量
                                    }
                                    else if ((!isXZ && isGH))
                                    {
                                        gh = "设施农用地";
                                        dic_increase[gh] += mj;    // 增加量
                                    }
                                }
                            }
                        }

                        // 属性映射现状用地\规划用地
                        string cySheet = $@"{targetExcel}\村域$";
                        ExcelTool.AttributeMapperDouble(cySheet, 11, 5, dic_xz, 4);
                        ExcelTool.AttributeMapperDouble(cySheet, 11, 9, dic_gh, 4);
                        // 属性映射增减量
                        ExcelTool.AttributeMapperDouble(cySheet, 11, 6, dic_increase, 4);
                        ExcelTool.AttributeMapperDouble(cySheet, 11, 7, dic_reduce, 4);


                        pw.AddMessageMiddle(10, $"生成村庄表", Brushes.Gray);
                        // 提取村庄的图斑
                        string tb_cz = $@"{tb}_cz";
                        Arcpy.TableSelect(tb, tb_cz, $"{czcField_xz} IN ('203', '203A') Or {czcField_gh} IN ('203', '203A')");

                        // 汇总指标
                        List<string> czList_xz = new List<string>() { gn_xz4 };
                        List<string> czList_gh = new List<string>() { gn_gh4 };
                        Dictionary<string, double> dic_cz_xz = ComboTool.MultiStatisticsToDic(tb_cz, $"Sum_{mjField}", czList_xz, "", unit_xs);
                        Dictionary<string, double> dic_cz_gh = ComboTool.MultiStatisticsToDic(tb_cz, $"Sum_{mjField}", czList_gh, "", unit_xs);
                        // 统计用地变化量
                        // 初始化用地集合
                        List<string> ydList_cz = new List<string>()
                        {
                            "一类农村宅基地","二类农村宅基地","农村社区服务设施用地","公共管理与公共服务用地","商业服务业用地",
                            "一类工业用地","二类工业用地","三类工业用地","采矿用地","一类物流仓储用地","二类物流仓储用地","三类物流仓储用地",
                            "储备库用地","交通运输用地","公用设施用地","绿地与开敞空间用地","留白用地","203内的其他用地",
                        };
                        Dictionary<string, double> dic_cz_increase = new Dictionary<string, double>();   // 增加量
                        Dictionary<string, double> dic_cz_reduce = new Dictionary<string, double>();     // 减少量
                        foreach (var yd in ydList_cz)
                        {
                            dic_cz_increase.Add(yd, 0);
                            dic_cz_reduce.Add(yd, 0);
                        }
                        // 遍历表格，更新增减量指标
                        Table table_cz = tb_cz.TargetTable();
                        // 逐行找出错误
                        using RowCursor rowCursor_cz = table_cz.Search();
                        while (rowCursor_cz.MoveNext())
                        {
                            using Row row = rowCursor_cz.Current;
                            // 获取value
                            double mj = double.Parse(row["SUM_jsmj"].ToString()) / unit_xs;
                            var va_xz = row[gn_xz4];
                            var va_gh = row[gn_gh4];
                            if (va_xz != null && va_gh != null)
                            {
                                // 获取现状规划名称
                                string xz = va_xz.ToString();
                                string gh = va_gh.ToString();
                                // 分析增减
                                if (xz != gh)
                                {
                                    if (xz != "")
                                    {
                                        dic_cz_reduce[xz] += mj;    // 减少量
                                    }
                                    if (gh != "")
                                    {
                                        dic_cz_increase[gh] += mj;    // 增加量
                                    }
                                }
                            }
                        }

                        // 属性映射现状用地\规划用地
                        string cySheet_cz = $@"{targetExcel}\村庄$";
                        ExcelTool.AttributeMapperDouble(cySheet_cz, 10, 4, dic_cz_xz, 4);
                        ExcelTool.AttributeMapperDouble(cySheet_cz, 10, 8, dic_cz_gh, 4);
                        // 属性映射增减量
                        ExcelTool.AttributeMapperDouble(cySheet_cz, 10, 5, dic_cz_increase, 4);
                        ExcelTool.AttributeMapperDouble(cySheet_cz, 10, 6, dic_cz_reduce, 4);
                        // 写入村庄表的增减总量
                        double increase = double.Parse(ExcelTool.GetCellFromExcel(cySheet, 16, 6));
                        double reduce = double.Parse(ExcelTool.GetCellFromExcel(cySheet, 16, 7));
                        ExcelTool.WriteCell(cySheet_cz, 22, 5, increase);
                        ExcelTool.WriteCell(cySheet_cz, 22, 6, reduce);

                        // 删除指定列
                        ExcelTool.DelectColSimple(cySheet, new List<int>() { 11 });
                        ExcelTool.DelectColSimple(cySheet_cz, new List<int>() { 10 });

                        Arcpy.Delect(tb);
                        Arcpy.Delect(tb_cz);
                    }
                    // 删除中间数据
                    Arcpy.Delect(staTable);
                    Arcpy.DeleteField(in_fc, mjField);
                    File.Delete(targetMapper);
                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }
    }
}
