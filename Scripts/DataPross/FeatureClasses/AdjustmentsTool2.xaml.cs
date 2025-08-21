using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
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

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    /// <summary>
    /// Interaction logic for AdjustmentsTool2.xaml
    /// </summary>
    public partial class AdjustmentsTool2 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public AdjustmentsTool2()
        {
            InitializeComponent();

            combox_areaType.Items.Add("投影面积");
            combox_areaType.Items.Add("图斑面积");
            combox_areaType.SelectedIndex = 1;

            // 初始化combox
            combox_digit.Items.Add("1");
            combox_digit.Items.Add("2");
            combox_digit.Items.Add("3");
            combox_digit.Items.Add("4");
            combox_digit.Items.Add("5");
            combox_digit.Items.Add("6");
            combox_digit.SelectedIndex = 1;
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "面积平差工具_按总面积";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_field);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string fc_field = combox_field.ComboxText();
                int digit = int.Parse(combox_digit.Text);
                double totalArea= double.Parse( textTotalArea.Text);
                string area_type = combox_areaType.Text[..2];

                // 默认数据库位置
                var gdb_path = Project.Current.DefaultGeodatabasePath;
                // 工程默认文件夹位置
                string folder_path = Project.Current.HomeFolderPath;

                // 判断参数是否选择完全
                if (fc_path == "" || fc_field == "" || combox_digit.Text == "")
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
                    pw.AddMessageStart("平差计算");
                    // 获取原始字段
                    List<string> fieldList = GisTool.GetFieldsNameFromTarget(fc_path);

                    string resultLayer = "";

                    resultLayer = ComboTool.AdjustmentByArea(fc_path, fc_field,  area_type,  digit, totalArea, gdb_path + @"\Adjustment");

                    pw.AddMessageMiddle(20, "覆盖图层", Brushes.Gray);
                    // 返回覆盖图层
                    string fcFullPath = fc_path.TargetLayerPath();
                    Arcpy.CopyFeatures(resultLayer, fcFullPath, true);

                    pw.AddMessageMiddle(20, "删除中间字段", Brushes.Gray);
                    // 删除中间字段
                    string fields = "";
                    foreach (var field in fieldList)
                    {
                        fields += field + ";";
                    }
                    string re = fields.Substring(0, fields.Length - 1);
                    Arcpy.DeleteField(fcFullPath, re, "KEEP_FIELDS");
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
            string url = "https://blog.csdn.net/xcc34452366/article/details/135822374?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

    }
}
