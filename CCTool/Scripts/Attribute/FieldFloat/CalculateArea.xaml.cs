using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using NPOI.Util;
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

namespace CCTool.Scripts.Attribute.FieldFloat
{
    /// <summary>
    /// Interaction logic for CalculateArea.xaml
    /// </summary>
    public partial class CalculateArea : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public CalculateArea()
        {
            InitializeComponent();
            // 初始化combox
            combox_unit.Items.Add("平方米");
            combox_unit.Items.Add("公顷");
            combox_unit.Items.Add("平方公里");
            combox_unit.Items.Add("亩");
            combox_unit.SelectedIndex = 0;

            combox_areaType.Items.Add("投影面积");
            combox_areaType.Items.Add("椭球面积");
            combox_areaType.SelectedIndex = 0;

            combox_digit.Items.Add("1");
            combox_digit.Items.Add("2");
            combox_digit.Items.Add("3");
            combox_digit.Items.Add("4");
            combox_digit.Items.Add("5");
            combox_digit.Items.Add("6");
            combox_digit.SelectedIndex = 1;
            
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void combox_field_DropOpen(object sender, EventArgs e)
        {
            // 获取参数
            string fc_path = combox_fc.ComboxText();
            UITool.AddFloatFieldsToComboxPlus(fc_path, combox_field);
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "计算面积";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string area_field = combox_field.ComboxText();
                string area_type = combox_areaType.Text;
                string unit = combox_unit.Text;
                int digit = int.Parse(combox_digit.Text);

                // 判断参数是否选择完全
                if (fc_path == "" || area_field == "")
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
                    List<string> errs = CheckData(fc_path);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    // 获取图层的所有字段
                    List<string> fields = GisTool.GetFieldsNameFromTarget(fc_path, "float");
                    pw.AddMessageMiddle(10,"添加字段", Brushes.Gray);
                    // 添加字段
                    if (!fields.Contains(area_field))
                    {
                        Arcpy.AddField(fc_path, area_field, "DOUBLE");
                    }
                    pw.AddMessageMiddle(20, "分析面积类型、单位等参数", Brushes.Gray);
                    // 单位系数
                    double xs = unit switch
                    {
                        "平方米" => 1,
                        "公顷" => 10000,
                        "平方公里" => 1000000,
                        "亩" => 666.66667,
                        _ => 1,
                    };
                    // 面积类型
                    string areaType = area_type switch
                    {
                        "投影面积"=>"area",
                        "椭球面积" => "geodesicarea",
                        _ => "",
                    };
                    pw.AddMessageMiddle(30, "计算面积", Brushes.Gray);
                    // 计算面积
                    Arcpy.CalculateField(fc_path, area_field, $"round(!shape.{areaType}!/{xs},{digit})");
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
            string url = "https://mp.weixin.qq.com/s/KqfbFAG3Tz6YUXj9lJxGAA";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string in_data)
        {
            List<string> result = new List<string>();


            return result;
        }
    }
}
