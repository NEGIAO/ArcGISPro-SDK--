using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
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

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    /// <summary>
    /// Interaction logic for ClosePolyline.xaml
    /// </summary>
    public partial class ClosePolyline : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "ClosePolyline";

        public ClosePolyline()
        {
            InitializeComponent();

            // 如果刚开始注册表没有值，就赋一个默认值
            string minDistance = BaseTool.ReadValueFromReg(toolSet, "minDistance");
            if (minDistance == "")
            {
                text_distance.Text = "10";
            }
            else
            {
                text_distance.Text = minDistance;
            }
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "平行线两端闭合";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc, "Polyline");
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string lyName = combox_fc.ComboxText();
                double minDistance = text_distance.Text.ToDouble();

                // 默认数据库位置
                var gdb_path = Project.Current.DefaultGeodatabasePath;
                // 工程默认文件夹位置
                string folder_path = Project.Current.HomeFolderPath;

                // 判断参数是否选择完全
                if (lyName == "" || minDistance == 0)
                {
                    MessageBox.Show("有必选参数为空，或参数填写错误！！！");
                    return;
                }

                BaseTool.WriteValueToReg(toolSet, "minDistance", minDistance);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageMiddle(10, "复制一个要素");
                    // 复制一个新的图层【去掉Z值】
                    string newLayer = @$"{gdb_path}\{lyName}_closed";
                    string fcName = $"{lyName}_closed";
                    Arcpy.CopyFeatures(lyName, newLayer, true, "Disabled");

                    FeatureLayer lineLayer = fcName.TargetFeatureLayer();

                    pw.AddMessageMiddle(20, "不同线段之间就近相连");
                    // 不同线段之间就近相连
                    GeometryTool.MergeNearLine(lineLayer, minDistance);
                    pw.AddMessageMiddle(20, "同线段首尾相连");
                    // 同线段首尾相连
                    GeometryTool.CloseLine(gdb_path, fcName, minDistance);

                    // 刷新地图视图
                    MapView.Active.ZoomInFixed();

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
            string url = "https://blog.csdn.net/xcc34452366/article/details/141418991";
            UITool.Link2Web(url);
        }
    }
}
