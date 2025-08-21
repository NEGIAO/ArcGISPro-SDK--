using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.Formula.Functions;
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

namespace CCTool.Scripts.LayerExport
{
    /// <summary>
    /// Interaction logic for Export2CADPlus.xaml
    /// </summary>
    public partial class Export2CADPlus : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public Export2CADPlus()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "要素图层导出CAD";

        private async void btn_go_click(object sender, RoutedEventArgs e)
        {
            // 获取指标
            string txtFront = txt_front.Text;
            string txtMid = txt_mid.Text;
            string txtBack = txt_back.Text;

            string field01 = combox_field01.ComboxText();
            string field02 = combox_field02.ComboxText();

            string cadPath = textCADPath.Text;

            // 如果什么都没填写，则直接导出
            bool isCal = txtFront == "" && txtMid == "" && txtBack == "" && field01 == "" && field02 == "";

            // 参数转义
            string field_01 = "''";
            string field_02 = "''";
            if (field01 != "")
            {
                field_01 = $"str(!{field01}!)";
            }
            if (field02 != "")
            {
                field_02 = $"str(!{field02}!)";
            }

            // 打开进度框
            ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
            pw.AddMessageTitle(tool_name);

            Close();

            await QueuedTask.Run(() =>
            {
                // 获取图层
                FeatureLayer ly = MapView.Active.GetSelectedLayers().FirstOrDefault() as FeatureLayer;

                if (!isCal)
                {
                    pw.AddMessageStart("添加Layer字段");
                    // 添加字段
                    string fieldName = "Layer";
                    if (!GisTool.IsHaveFieldInTarget(ly, fieldName))
                    {
                        Arcpy.AddField(ly, fieldName, "TEXT");
                    }
                    pw.AddMessageMiddle(10, "计算Layer字段");
                    // 计算字段
                    Arcpy.CalculateField(ly, fieldName, $"'{txtFront}' + {field_01} +'{txtMid}' +{field_02} +'{txtBack}'");
                }
                pw.AddMessageMiddle(10, "导出CAD");
                // 导出CAD
                Arcpy.ExportCAD(ly, cadPath);
            });

            pw.AddMessageEnd();
        }

        private void btn_help_click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/141057275";
            UITool.Link2Web(url);
        }

        private void combox_field01_DropOpen(object sender, EventArgs e)
        {
            FeatureLayer ly = MapView.Active.GetSelectedLayers().FirstOrDefault() as FeatureLayer;
            UITool.AddFieldsToComboxPlus(ly, combox_field01);
        }

        private void combox_field02_DropOpen(object sender, EventArgs e)
        {
            FeatureLayer ly = MapView.Active.GetSelectedLayers().FirstOrDefault() as FeatureLayer;
            UITool.AddFieldsToComboxPlus(ly, combox_field02);
        }

        private void openTableButton_Click(object sender, RoutedEventArgs e)
        {
            textCADPath.Text = UITool.SaveDialogCAD();
        }
    }
}
