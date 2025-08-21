using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
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

namespace CCTool.Scripts.CusTool2
{
    /// <summary>
    /// Interaction logic for StatisticsYDYH2.xaml
    /// </summary>
    public partial class StatisticsYDYH2 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public StatisticsYDYH2()
        {
            InitializeComponent();
            Init();       // 初始化
        }

        // 初始化
        private void Init()
        {
            combox_unit.Items.Add("平方米");
            combox_unit.Items.Add("公顷");
            combox_unit.Items.Add("平方公里");
            combox_unit.Items.Add("亩");
            combox_unit.SelectedIndex = 1;

            combox_digit.Items.Add("1");
            combox_digit.Items.Add("2");
            combox_digit.Items.Add("3");
            combox_digit.Items.Add("4");
            combox_digit.Items.Add("5");
            combox_digit.Items.Add("6");
            combox_digit.SelectedIndex = 3;
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "用地用海指标汇总(特殊定制)";

        // 点击打开按钮，选择输出的Excel文件位置
        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            // 打开Excel文件
            string path = UITool.SaveDialogExcel();
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
                var init_gdb = Project.Current.DefaultGeodatabasePath;
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string field_bm = combox_bmField.ComboxText();
                string excel_path = textExcelPath.Text;

                string unit = combox_unit.Text;
                int digit = int.Parse(combox_digit.Text);

                string areaField = combox_areaField.ComboxText();

                List<string> bmList = new List<string>() { field_bm };

                // 判断参数是否选择完全
                if (fc_path == "" || field_bm == "" || excel_path == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                // 单位系数设置
                double unit_xs = unit switch
                {
                    "平方米" => 1,
                    "公顷" => 10000,
                    "平方公里" => 1000000,
                    "亩" => 666.66667,
                    _ => 1,
                };

                Close();
                await QueuedTask.Run(() =>
                {
                    // 表名
                    string excelName = "规划汇总表(特殊定制)";
                    // 复制嵌入资源中的Excel文件
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.{excelName}.xlsx", excel_path);
                    string excel_sheet = excel_path + @"\Sheet1$";
                    // 设置小数位数
                    ExcelTool.SetDigit(excel_sheet, new List<int>() { 5,6 }, 4, digit);

                    pw.AddMessageStart("初步汇总指标");
                    // 初步汇总指标
                    Dictionary<string, double> dic = ComboTool.MultiStatisticsToDic(fc_path, areaField, bmList, "总用地", unit_xs);
                    
                    pw.AddMessageMiddle(10, "汇总一级类指标");
                    // 汇总一级类指标
                    string fc2 = init_gdb + @"\fc2";
                    string newField = "NewDM";
                    Arcpy.CopyFeatures(fc_path, fc2);
                    Arcpy.AddField(fc2, newField, "TEXT");
                    ComboTool.AttributeMapper(fc2, field_bm, newField, excel_path + @"\Sheet2$");
                    Dictionary<string, double> dic2 = ComboTool.MultiStatisticsToDic(fc2, areaField, new List<string>() { newField }, "总用地", unit_xs);

                    pw.AddMessageMiddle(10, "汇总二级类指标");
                    // 汇总二级类指标
                    ComboTool.AttributeMapper(fc2, field_bm, newField, excel_path + @"\Sheet3$");
                    Dictionary<string, double> dic3 = ComboTool.MultiStatisticsToDic(fc2, areaField, new List<string>() { newField }, "总用地", unit_xs);

                    pw.AddMessageMiddle(10, "汇总非建和建设用地");
                    // 汇总二级类指标
                    ComboTool.AttributeMapper(fc2, field_bm, newField, excel_path + @"\Sheet4$");
                    Dictionary<string, double> dic4 = ComboTool.MultiStatisticsToDic(fc2, areaField, new List<string>() { newField }, "总用地", unit_xs);

                    pw.AddMessageMiddle(10, "指标写入Excel");
                    // 属性映射
                    ExcelTool.AttributeMapperDouble(excel_sheet, 8, 5, dic, 2);
                    ExcelTool.AttributeMapperDouble(excel_sheet, 8, 5, dic2, 2);
                    ExcelTool.AttributeMapperDouble(excel_sheet, 8, 5, dic3, 2);
                    ExcelTool.AttributeMapperDouble(excel_sheet, 8, 5, dic4, 2);
                    // 删除0值行
                    ExcelTool.DeleteNullRow(excel_sheet, 5, 2);
                    // 删除指定列
                    ExcelTool.DeleteCol(excel_sheet, new List<int>() { 8 });
                    // 改Excel中的单位
                    ExcelTool.WriteCell(excel_path, 1, 5, $"面积({unit})");
                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        private void combox_areaField_DropDown(object sender, EventArgs e)
        {
            string fc_path = combox_fc.ComboxText();
            UITool.AddAllFloatFieldsToComboxPlus(fc_path, combox_areaField);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "";
            UITool.Link2Web(url);
        }

        private void combox_fc_Closed(object sender, EventArgs e)
        {
            try
            {
                // 填写输出路径
                textExcelPath.Text = Project.Current.HomeFolderPath + @"\用地用海指标汇总表(特殊).xlsx";
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

    }
}
