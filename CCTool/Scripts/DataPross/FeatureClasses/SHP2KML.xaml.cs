using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using Org.BouncyCastle.Asn1.Cms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

namespace CCTool.Scripts.DataPross.SHP
{
    /// <summary>
    /// Interaction logic for SHP2KML.xaml
    /// </summary>
    public partial class SHP2KML : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "SHP2KML2";

        public SHP2KML()
        {
            InitializeComponent();

            // 初始化参数选项
            folderPath.Text = BaseTool.ReadValueFromReg(toolSet, "folder_path");

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "要素图层按字段导出KMZ";

        private void openForderButton_Click(object sender, RoutedEventArgs e)
        {
            folderPath.Text = UITool.OpenDialogFolder();
        }

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            string fc = combox_fc.ComboxText();
            UITool.AddFieldsToComboxPlus(fc, combox_field);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc, "All");
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string folder_path = folderPath.Text;
                string fc = combox_fc.ComboxText();
                string shot_field = combox_field.ComboxText();

                // 判断参数是否选择完全
                if (folder_path == "" || fc == "" || shot_field == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 初始化参数选项
                BaseTool.WriteValueToReg(toolSet, "folder_path", folder_path);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                // 异步执行
                await QueuedTask.Run(() =>
                {
                    // 获取图层
                    FeatureLayer featureLayer = fc.TargetFeatureLayer();

                    // 编码字段值
                    List<string> fieldValues = GisTool.GetFieldValuesFromPath(fc, shot_field);

                    // 按编码导出KMZ
                    foreach (string fieldValue in fieldValues)
                    {
                        pw.AddMessageMiddle(90 / fieldValues.Count, $"导出_{fieldValue}");

                        var queryFilter = new QueryFilter();
                        queryFilter.WhereClause = $"{shot_field} = '{fieldValue}'";
                        // 图层选择
                        featureLayer.Select(queryFilter);

                        // 导出
                        Arcpy.LayerToKML(fc, $@"{folder_path}\k_{fieldValue}.kmz");
                    }

                    // 取消选择
                    MapCtlTool.UnSelectAllFeature(fc);

                    pw.AddMessageEnd();
                });

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135840446?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

    }
}
