using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Windows;
using NPOI.OpenXmlFormats.Vml;
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

namespace CCTool.Scripts.UI.SD
{
    /// <summary>
    /// Interaction logic for SDStatistic1.xaml
    /// </summary>
    public partial class SDStatistic1 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {

        // 工具设置标签
        readonly string toolSet = "SDStatistic1";

        public SDStatistic1()
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

            // 初始化参数选项
            textTablePath.Text = BaseTool.ReadValueFromReg(toolSet, "excelPath");
            string siAdj = BaseTool.ReadValueFromReg(toolSet, "isAdj");
            string isVg = BaseTool.ReadValueFromReg(toolSet, "isVg");

            checkBox_adj.IsChecked = siAdj == "True";
            checkBox_vg.IsChecked = isVg == "True";
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "2、土地利用现状一级分类面积汇总表";


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
                string fc = combox_fc.ComboxText();
                string areaType = combox_area.Text[..2];
                string unit = combox_unit.Text;
                _ = int.TryParse(combox_digit.Text, out int digit);
                bool isAdj = (bool)checkBox_adj.IsChecked;

                bool isVg = (bool)checkBox_vg.IsChecked;

                string zoom = combox_fc_area.ComboxText();
                // 分区字段
                string oldNameField = combox_field_area.ComboxText();
                // 由于分区字段可能和三调图层的字段同名，因此需要重新命名
                string nameField = "FQBSMC";

                string excelPath = textTablePath.Text;
                // 默认数据库位置
                var defGDB = Project.Current.DefaultGeodatabasePath;
                // 工程默认文件夹位置
                string defFolder = Project.Current.HomeFolderPath;

                // 判断参数是否选择完全
                if (fc == "" || excelPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }
                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "excelPath", excelPath);
                BaseTool.WriteValueToReg(toolSet, "isAdj", isAdj.ToString());
                BaseTool.WriteValueToReg(toolSet, "isVg", isVg.ToString());

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(fc, zoom, oldNameField);
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
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.三调规程.土地利用现状一级分类面积汇总表.xlsx", excelPath);
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.三调用地自转换.xlsx", defFolder + @"\三调用地自转换.xlsx");

                    pw.AddMessageMiddle(10, "目标图层预处理");

                    string targetFc = @$"{defGDB}\targetFc";
                    string zoomDissolve = @$"{defGDB}\zoomDissolve";
                    string identity = @$"{defGDB}\identity";
                    if (zoom != "")  // 如果分地块统计，进行标识
                    {
                        // 先融合分地块
                        Arcpy.Dissolve(zoom, zoomDissolve, oldNameField);
                        // 重命名分区字段
                        Arcpy.AlterField(zoomDissolve, oldNameField, nameField, nameField);
                        // 标识
                        Arcpy.Identity(fc, zoomDissolve, identity);
                        // 选择
                        Arcpy.Select(identity, targetFc, $"{nameField}  <> ''");
                    }
                    else  // 如果没有分地块统计，添加一个字段进行标记
                    {
                        Arcpy.CopyFeatures(fc, targetFc);
                        Arcpy.AddField(targetFc, nameField, "TEXT");
                        // 图层名简化
                        string layerName = fc[(fc.LastIndexOf('\\') + 1)..];
                        Arcpy.CalculateField(targetFc, nameField, $"'{layerName}'");
                        // 分区域
                        Arcpy.Dissolve(targetFc, zoomDissolve, nameField);
                    }

                    // 平差计算
                    if (isAdj)
                    {
                        pw.AddMessageMiddle(20, $"平差计算");
                        targetFc = ComboTool.AreaAdjustment(zoomDissolve, targetFc, nameField, areaType, unit, digit);
                    }
                    else
                    {
                        targetFc = ComboTool.AreaAdjustmentNot(targetFc, areaType, unit, digit);
                    }

                    pw.AddMessageMiddle(10, "用地名称转为统计分类字段");
                    // 添加一个类一、二级字段
                    Arcpy.AddField(targetFc, "mc_1", "TEXT");

                    // 用地编码转换
                    ComboTool.AttributeMapper(targetFc, "DLMC", "mc_1", defFolder + @"\三调用地自转换.xlsx\一级$");

                    // 如果不分村统计，就把【ZLDWDM;ZLDWMC】字段统一一下
                    if (!isVg)
                    {
                        Arcpy.CalculateField(targetFc, "ZLDWDM", "'XXX'");
                        Arcpy.CalculateField(targetFc, "ZLDWMC", "'XXX'");
                    }

                    // 计算扣除面积
                    string sdMJ = "SDMJ";
                    string tkMJ = "田坎面积";
                    Arcpy.AddField(targetFc, tkMJ, "DOUBLE");
                    Arcpy.CalculateField(targetFc, tkMJ, $"round(!{sdMJ}!* !KCXS!,{digit})");
                    Arcpy.CalculateField(targetFc, sdMJ, $"round(!{sdMJ}!- !{tkMJ}!,{digit})");

                    // 汇总
                    List<string> zoomNames = GisTool.GetFieldValuesFromPath(targetFc, nameField);
                    foreach (string zoomName in zoomNames)
                    {
                        pw.AddMessageMiddle(10, $"处理表格：{zoomName}");

                        // 新建sheet
                        ExcelTool.CopySheet(excelPath, "sheet1", zoomName);
                        string new_sheet_path = excelPath + @$"\{zoomName}$";

                        // 分村庄
                        List<string> vgNames = GisTool.GetFieldValuesFromPath(targetFc, "ZLDWMC");
                        //  村庄名称和代码
                        Dictionary<string, string> dmmcDict = GisTool.GetDictFromPath(targetFc, "ZLDWMC", "ZLDWDM");

                        // 处理每个行政单位
                        int start_row = 6;
                        List<int> row_list = new List<int>();
                        foreach (string vgName in vgNames)
                        {
                            // 获取行政代码
                            string vgCode = dmmcDict[vgName];

                            pw.AddMessageMiddle(10, $"      {vgName}__汇总用地指标", Brushes.Gray);

                            // 收集扣除田坎的三调用地面积， 以及田坎面积
                            List<string> mcFields = new List<string> { "mc_1"};
                            string sql = $"{nameField} = '{zoomName}' AND ZLDWMC = '{vgName}'";
                            Dictionary<string, double> sdDict = ComboTool.StatisticsPlus(targetFc, mcFields, sdMJ, "国土调查总面积", 1, sql);
                            Dictionary<string, double> tkDict = ComboTool.StatisticsPlus(targetFc, "DLMC", tkMJ, "合计", 1, sql);
                            // 如果没有地类，就跳过
                            if (!sdDict.ContainsKey("国土调查总面积"))
                            {
                                continue;
                            }

                            //  Dict加入田坎面积
                            double tk = tkDict.ContainsKey("合计") ? tkDict["合计"] : 0;
                            sdDict.Add("田坎", tk);

                            sdDict["国土调查总面积"] += tk;

                            if (sdDict.ContainsKey("其他土地"))
                            {
                                sdDict["其他土地"] += tk;
                            }
                            else
                            {
                                sdDict.Add("其他土地", tk);
                            }

                            // 复制行
                            ExcelTool.CopyRows(new_sheet_path, 5, start_row);
                            // 记录写入行数的列表
                            row_list.Add(start_row);

                            // 属性映射到Excel
                            ExcelTool.AttributeMapperColDouble(new_sheet_path, 1, start_row, sdDict);

                            // 写入名称和代码
                            ExcelTool.WriteCell(new_sheet_path, start_row, 0, vgName);
                            ExcelTool.WriteCell(new_sheet_path, start_row, 1, vgCode);

                            // 进入下一个行政单位
                            start_row++;
                        }

                        // 更新面积标识
                        ExcelTool.WriteCell(new_sheet_path, 2, 0, $"面积单位：{unit}");

                        // 删除空列
                        ExcelTool.DeleteNullCol(new_sheet_path, row_list, 2);
                        // 删除行
                        ExcelTool.DeleteRow(new_sheet_path, new List<int>() { 5, 1 });

                    }

                    // 删除sheet1
                    ExcelTool.DeleteSheet(excelPath, "Sheet1");

                    // 删除中间数据
                    Arcpy.Delect(zoomDissolve);
                    Arcpy.Delect(identity);
                    Arcpy.Delect(targetFc);
                    Arcpy.Delect(@$"{defGDB}\SortFc");

                    File.Delete(defFolder + @"\三调用地自转换.xlsx");

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

        }

        private void combox_field_area_DropDown(object sender, EventArgs e)
        {
            string area_fc = combox_fc_area.ComboxText();
            UITool.AddTextFieldsToComboxPlus(area_fc, combox_field_area);
        }
       
        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135698789?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string sd, string area, string in_field)
        {
            List<string> result = new List<string>();
            // 检查三调
            if (sd != "")
            {
                // 检查字段值是否符合要求
                string result_value = CheckTool.CheckFieldValue(sd, "DLMC", GlobalData.yd_sd);
                if (result_value != "")
                {
                    result.Add(result_value);
                }
                else
                {
                    // 检查字段值是否为空【KCXS】
                    string fieldEmptyResult = CheckTool.CheckFieldValueEmpty(sd, "KCXS", "DLMC IN ('旱地', '水田', '水浇地')");
                    if (fieldEmptyResult != "")
                    {
                        result.Add(fieldEmptyResult);
                    }
                }

                // 检查字段值是否为空【ZLDWDM,ZLDWMC】
                string fieldEmptyResult2 = CheckTool.CheckFieldValueSpace(sd, new List<string>() { "ZLDWDM", "ZLDWMC" });
                if (fieldEmptyResult2 != "")
                {
                    result.Add(fieldEmptyResult2);
                }
            }

            // 检查是否选了分地块但没选择字段
            if (area != "" && in_field == "")
            {
                result.Add("如果启用了分地块，请选择分地块名称字段。");
            }

            // 检查分地块
            if (area != "" && in_field != "")
            {
                // 检查字段值是否为空
                string fieldEmptyResult = CheckTool.CheckFieldValueSpace(area, in_field);
                if (fieldEmptyResult != "")
                {
                    result.Add(fieldEmptyResult);
                }

                // 检查一下分地块和三调的坐标系是否一致
                string srResult = CheckTool.CheckSpatialReference(sd, area);

                if (srResult != "")
                {
                    result.Add(srResult);
                }

            }

            // 检查是否正常提取Excel
            string result_excel = CheckTool.CheckExcelPick();
            if (result_excel != "")
            {
                result.Add(result_excel);
            }

            return result;
        }

    }
}
