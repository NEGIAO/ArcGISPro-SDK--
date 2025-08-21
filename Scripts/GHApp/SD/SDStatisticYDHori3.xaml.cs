using ArcGIS.Core.Data;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Mapping.Locate;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Windows;
using CCTool.Scripts.UI.ProWindow;
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
using Row = ArcGIS.Core.Data.Row;
using Table = ArcGIS.Core.Data.Table;


namespace CCTool.Scripts.GHApp.SD
{
    /// <summary>
    /// Interaction logic for SDStatisticYDHori3.xaml
    /// </summary>
    public partial class SDStatisticYDHori3 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public SDStatisticYDHori3()
        {
            InitializeComponent();
            // 初始化combox
            combox_unit.Items.Add("平方米");
            combox_unit.Items.Add("公顷");
            combox_unit.Items.Add("平方公里");
            combox_unit.Items.Add("亩");
            combox_unit.SelectedIndex = 0;

            combox_area.Items.Add("投影面积");
            combox_area.Items.Add("图斑面积");
            combox_area.SelectedIndex = 0;

            combox_digit.Items.Add("1");
            combox_digit.Items.Add("2");
            combox_digit.Items.Add("3");
            combox_digit.Items.Add("4");
            combox_digit.Items.Add("5");
            combox_digit.Items.Add("6");
            combox_digit.SelectedIndex = 1;

            combox_field_area.IsEnabled = false;

            // 隐藏错误标记
            errorButton_sd.Visibility = Visibility.Hidden;
            errorButton_area.Visibility = Visibility.Hidden;
            errorButton_areaField.Visibility = Visibility.Hidden;
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "1、土地利用现状分类面积汇总表(含三级类)";

        // 定义一个错误信息框
        private MsgWindow msgWindow = null;
        // 三调图层检查
        Dictionary<string, SolidColorBrush> error_sd = new Dictionary<string, SolidColorBrush>();
        Dictionary<string, SolidColorBrush> error_area = new Dictionary<string, SolidColorBrush>();
        Dictionary<string, SolidColorBrush> error_areaField = new Dictionary<string, SolidColorBrush>();

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void openTableButton_Click(object sender, RoutedEventArgs e)
        {
            textTablePath.Text = UITool.SaveDialogExcel();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string area_type = combox_area.Text[..2]; // 去掉“面积”，以防shp故障
                string unit = combox_unit.Text;
                int digit = int.Parse(combox_digit.Text);
                bool isAdj = (bool)checkBox_adj.IsChecked;

                string fc_area_path = combox_fc_area.ComboxText();
                string name_field = combox_field_area.ComboxText();

                string excel_path = textTablePath.Text;
                // 默认数据库位置
                var gdb_path = Project.Current.DefaultGeodatabasePath;
                // 工程默认文件夹位置
                string folder_path = Project.Current.HomeFolderPath;

                // 判断参数是否选择完全
                if (fc_path == "" || excel_path == "")
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

                    // 获取图层名
                    string single_layer_path = fc_path.GetLayerSingleName();

                    // 复制嵌入资源中的Excel文件
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.三调规程.土地利用现状分类面积汇总表(含三级类).xlsx", excel_path);
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.三调用地自转换.xlsx", folder_path + @"\三调用地自转换.xlsx");

                    // 放到新建的数据库中
                    string filePath = $@"{folder_path}\CC配置文件(勿删)";
                    if (!Directory.Exists(filePath)) { Directory.CreateDirectory(filePath); }
                    string gdbName = "临时数据库";
                    Arcpy.CreateFileGDB(filePath, gdbName);
                    string new_gdb = $@"{filePath}\{gdbName}.gdb";
                    // 清理一下
                    GisTool.ClearGDBItem(new_gdb);

                    if (fc_area_path == "")  // 如果没有分地块统计
                    {
                        Arcpy.Dissolve(fc_path, $@"{new_gdb}\T{single_layer_path}");
                    }

                    else   // 如果分地块统计
                    {
                        // 先融合
                        Arcpy.Dissolve(fc_area_path, gdb_path + @"\dissolveFC", name_field);
                        // 按属性分割
                        Arcpy.SplitByAttributes(gdb_path + @"\dissolveFC", new_gdb, name_field);
                    }

                    // 收集分割后的地块
                    List<string> list_area = new_gdb.GetFeatureClassAndTablePath();

                    foreach (var area in list_area)
                    {
                        // 裁剪新用地
                        string area_name = area[(area.LastIndexOf(@"\") + 1)..];
                        string adj_fc = $@"{new_gdb}\{area_name}_结果";

                        pw.AddMessageMiddle(20, $"处理表格：{area_name}");

                        if (!isAdj)          // 如果不平差计算
                        {
                            ComboTool.AdjustmentNot(fc_path, area, adj_fc, area_type, unit, digit);
                        }
                        else          // 平差计算
                        {
                            pw.AddMessageMiddle(10, $"平差计算", Brushes.Gray);
                            ComboTool.Adjustment(fc_path, area, adj_fc, area_type, unit, digit);
                        }

                        pw.AddMessageMiddle(10, "用地名称转为统计分类字段", Brushes.Gray);
                        // 添加一个类一、二级字段
                        Arcpy.AddField(adj_fc, "mc_1", "TEXT");
                        Arcpy.AddField(adj_fc, "mc_2", "TEXT");

                        // 用地编码转换
                        ComboTool.AttributeMapper(adj_fc, "DLMC", "mc_1", folder_path + @"\三调用地自转换.xlsx\一级$");
                        ComboTool.AttributeMapper(adj_fc, "DLMC", "mc_2", folder_path + @"\三调用地自转换.xlsx\二级$");

                        // 新建sheet
                        ExcelTool.CopySheet(excel_path, "sheet1", area_name);
                        string new_sheet_path = excel_path + @$"\{area_name}$";

                        // 处理地块
                        // 要处理的要素名称
                        string area_name_detail = adj_fc[(adj_fc.LastIndexOf(@"\") + 1)..];

                        pw.AddMessageMiddle(10, $"{area_name_detail}__汇总用地指标", Brushes.Gray);

                        // 收集未扣除田坎的面积
                        string gdbBefore = gdb_path + @"\前";
                        Arcpy.Statistics(adj_fc, gdbBefore, area_type + " SUM", "");
                        double zmj_old = double.Parse(gdbBefore.TargetCellValue("SUM_" + area_type, "OBJECTID=1"));
                        // 计算扣除面积
                        Arcpy.CalculateField(adj_fc, area_type, $"round(!{area_type}!* (1-!KCXS!),{digit})");

                        // 收集扣除田坎的面积
                        string gdbAfter = gdb_path + @"\后";
                        Arcpy.Statistics(adj_fc, gdbAfter, area_type + " SUM", "");
                        double zmj_new = double.Parse(gdbAfter.TargetCellValue("SUM_" + area_type, "OBJECTID=1"));
                        double kcmj = Math.Round(zmj_old - zmj_new, digit);

                        // 汇总大、中类
                        List<string> list_bm = new List<string>() { "mc_1", "mc_2" };
                        string statistic_sd = gdb_path + @"\statistic_sd";
                        string statistic_sd2 = gdb_path + @"\statistic_sd2";
                        ComboTool.MultiStatistics(adj_fc, statistic_sd, area_type + " SUM", list_bm, "国土调查总面积", 4);

                        // 汇总三级类
                        List<string> list_xhdl = new List<string>() { "XHDL" };
                        ComboTool.MultiStatistics(adj_fc, statistic_sd2, area_type + " SUM", list_xhdl, "共计", 4);

                        // 统计田坎面积
                        if (kcmj != 0)
                        {
                            // 插入【田坎】行
                            ComboTool.UpdataRowToTable(statistic_sd, $"分组,田坎;SUM_{area_type},{kcmj}");
                            // 计算【其他土地、国土调查总面积】行
                            ComboTool.IncreRowValueToTable(statistic_sd, $"分组,其他土地;SUM_{area_type},{kcmj}");
                            ComboTool.IncreRowValueToTable(statistic_sd, $"分组,国土调查总面积;SUM_{area_type},{kcmj}");
                        }
                        // 再次确认小数位数
                        Arcpy.CalculateField(statistic_sd, $"SUM_{area_type}", $"round(!SUM_{area_type}!, {digit})");
                        // 将映射属性表中获取字典Dictionary
                        Dictionary<string, string> dict = GisTool.GetDictFromPath(statistic_sd, @"分组", "SUM_" + area_type);

                        Dictionary<string, double> dict2 = GisTool.GetDictFromPathDouble(statistic_sd2, @"分组", "SUM_" + area_type);

                        // 属性映射到Excel
                        ExcelTool.AttributeMapper(new_sheet_path, 7, 5, dict, 5);

                        ExcelTool.AttributeMapperDouble(new_sheet_path, 7, 5, dict2, 5);

                        pw.AddMessageMiddle(20, "删除空列", Brushes.Gray);

                        // 更新面积标识
                        ExcelTool.WriteCell(new_sheet_path, 2, 1, $"面积单位：{unit}");

                        // 删除空行
                        ExcelTool.DeleteNullRow(new_sheet_path, 5, 4);
                        // 删除参照列
                        ExcelTool.DeleteCol(new_sheet_path, 7);
                    }

                    // 删除sheet1
                    ExcelTool.DeleteSheet(excel_path, "Sheet1");

                    // 删除中间数据
                    Arcpy.Delect(gdb_path + @"\statistic_sd");
                    Arcpy.Delect(gdb_path + @"\dissolveFC");
                    Arcpy.Delect(gdb_path + @"\前");
                    Arcpy.Delect(gdb_path + @"\后");
                    File.Delete(folder_path + @"\三调用地自转换.xlsx");
                    // 收集分割后的地块
                    List<string> list = new_gdb.GetFeatureClassAndTablePath();
                    foreach (var item in list)
                    {
                        Arcpy.Delect(item);
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

        private void combox_fc_area_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc_area);
        }

        private void combox_fc_area_Closed(object sender, EventArgs e)
        {
            string fc_area = combox_fc_area.ComboxText();
            if (fc_area == "")
            {
                combox_field_area.IsEnabled = false;
            }
            else
            {
                combox_field_area.IsEnabled = true;
            }

            try
            {
                // 检查分地块图层
                CheckArea();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void combox_field_area_DropDown(object sender, EventArgs e)
        {
            string area_fc = combox_fc_area.ComboxText();
            UITool.AddTextFieldsToComboxPlus(area_fc, combox_field_area);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135688560?spm=1001.2014.3001.5502";
            UITool.Link2Web(url);
        }

        private void errorButton_sd_Click(object sender, RoutedEventArgs e)
        {
            // 打开错误信息框
            MsgWindow mw = UITool.OpenMsgWindow(msgWindow);
            foreach (var error in error_sd)
            {
                mw.AddMessage(error.Key, error.Value);
            }
        }

        private void combox_fc_Closed(object sender, EventArgs e)
        {
            try
            {
                // 填写输出路径
                textTablePath.Text = Project.Current.HomeFolderPath + @"\三调统计表_一级类.xlsx";
                // 检查三调图层
                CheckSD();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        // 检查三调图层
        public async void CheckSD()
        {
            // 隐藏错误标记
            errorButton_sd.Visibility = Visibility.Hidden;

            error_sd.Clear();

            // 三调图层
            string in_data = combox_fc.ComboxText();

            await QueuedTask.Run(() =>
            {
                if (in_data != "")
                {
                    // 检查字段值是否符合要求
                    string result_value = CheckTool.CheckFieldValue(in_data, "DLMC", GlobalData.yd_sd);
                    if (result_value != "")
                    {
                        error_sd.Add(result_value, Brushes.Red);
                    }
                    else
                    {
                        // 检查字段值是否为空【KCXS】
                        string fieldEmptyResult = CheckTool.CheckFieldValueEmpty(in_data, "KCXS", "DLMC IN ('旱地', '水田', '水浇地')");
                        if (fieldEmptyResult != "")
                        {
                            error_sd.Add(fieldEmptyResult, Brushes.Red);
                        }
                    }

                    // 检查字段值是否为空【ZLDWDM,ZLDWMC】
                    string fieldEmptyResult2 = CheckTool.CheckFieldValueSpace(in_data, new List<string>() { "ZLDWDM", "ZLDWMC" });
                    if (fieldEmptyResult2 != "")
                    {
                        error_sd.Add(fieldEmptyResult2, Brushes.Red);
                    }
                }
            });

            if (error_sd.Count > 0)
            {
                errorButton_sd.Visibility = Visibility.Visible;
            }
        }

        private void errorButton_area_Click(object sender, RoutedEventArgs e)
        {
            // 打开错误信息框
            MsgWindow mw = UITool.OpenMsgWindow(msgWindow);
            foreach (var error in error_area)
            {
                mw.AddMessage(error.Key, error.Value);
            }
        }

        private  void CheckArea()
        {
            // 隐藏错误标记
            errorButton_area.Visibility = Visibility.Hidden;

            error_area.Clear();

            // 分地块图层
            string in_data = combox_fc_area.ComboxText();


            if (error_area.Count > 0)
            {
                errorButton_area.Visibility = Visibility.Visible;
            }
        }

        private void errorButton_areaField_Click(object sender, RoutedEventArgs e)
        {
            // 打开错误信息框
            MsgWindow mw = UITool.OpenMsgWindow(msgWindow);
            foreach (var error in error_areaField)
            {
                mw.AddMessage(error.Key, error.Value);
            }
        }

        private async void CheckAreaField()
        {
            // 隐藏错误标记
            errorButton_areaField.Visibility = Visibility.Hidden;

            error_areaField.Clear();

            // 分地块图层
            string in_data = combox_fc_area.ComboxText();
            // 分地块名称字段
            string in_field = combox_field_area.ComboxText();

            await QueuedTask.Run(() =>
            {
                if (in_data != "" && in_field != "")
                {

                    // 检查字段值是否为空
                    string fieldEmptyResult = CheckTool.CheckFieldValueSpace(in_data, in_field);
                    if (fieldEmptyResult != "")
                    {
                        error_areaField.Add(fieldEmptyResult, Brushes.Red);
                    }
                }
            });

            if (error_areaField.Count > 0)
            {
                errorButton_areaField.Visibility = Visibility.Visible;
            }
        }

        private void combox_field_area_Closed(object sender, EventArgs e)
        {
            try
            {
                // 检查分地块名称字段
                CheckAreaField();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }
    }
}
