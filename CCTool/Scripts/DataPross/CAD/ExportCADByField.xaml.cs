using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
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

namespace CCTool.Scripts.DataPross.CAD
{
    /// <summary>
    /// Interaction logic for ExportCADByField.xaml
    /// </summary>
    public partial class ExportCADByField : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "ExportCADByField";

        public ExportCADByField()
        {
            InitializeComponent();

            // 初始化参数选项
            textCADFolder.Text = BaseTool.ReadValueFromReg(toolSet, "CADFolder");

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "按要素字段批量导出CAD";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc, "All");
        }


        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc = combox_fc.ComboxText();
                string field = combox_field.ComboxText();
                string CADFolder = textCADFolder.Text;

                // 判断参数是否选择完全
                if (fc == "" || field == "" || CADFolder == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }
                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "CADFolder", CADFolder);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                //Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(fc, field);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10, "导出CAD");
                    List<string> fields = GisTool.GetFieldValuesFromPath(fc, field);

                    FeatureLayer featureLayer = fc.TargetFeatureLayer();

                    foreach (string fd in fields)
                    {
                        pw.AddMessageMiddle(80/fields.Count, $"      导出CAD_{fd}", Brushes.Gray);
                        // 选择
                        QueryFilter queryFilter = new QueryFilter();
                        queryFilter.WhereClause = $"{field} = '{fd}' ";
                        featureLayer.Select(queryFilter);

                        // 导出
                        Arcpy.ExportCAD(fc, $@"{CADFolder}\{fd}.dwg");
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



        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147895772?spm=1001.2014.3001.5502";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string fc, string field)
        {
            List<string> result = new List<string>();

            // 检查字段值是否为空
            string fieldSpaceResult = CheckTool.CheckFieldValueSpace(fc, field);
            if (fieldSpaceResult != "")
            {
                result.Add(fieldSpaceResult);
            }
            return result;
        }

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field);
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textCADFolder.Text = UITool.OpenDialogFolder();
        }
    }
}
