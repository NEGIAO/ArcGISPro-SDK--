using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Windows;
using Microsoft.Office.Core;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.OpenXmlFormats.Vml;
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
    /// Interaction logic for DissolveAsField.xaml
    /// </summary>
    public partial class DissolveAsField : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "DissolveAsField";

        public DissolveAsField()
        {
            InitializeComponent();

            // 初始化其它参数选项
            textFeatureClassPath.Text = BaseTool.ReadValueFromReg(toolSet, "featureClass_path");
            text_mj.Text = BaseTool.ReadValueFromReg(toolSet, "minArea", "10");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "融合同类碎图斑";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                double minArea = text_mj.Text.ToDouble();
                string featureClass_path = textFeatureClassPath.Text;

                // 获取参数listbox
                List<string> fieldNames = UITool.GetCheckboxStringFromListBox(listbox_field);

                // 默认数据库位置
                var gdb_path = Project.Current.DefaultGeodatabasePath;
                // 工程默认文件夹位置
                string folder_path = Project.Current.HomeFolderPath;

                // 判断参数是否选择完全
                if (fc_path == "" || minArea <= 0 || fieldNames.Count == 0 || featureClass_path == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "featureClass_path", featureClass_path);
                BaseTool.WriteValueToReg(toolSet, "minArea", minArea);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(featureClass_path);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    // 复制要素
                    string targetFC = $@"{gdb_path}\dissolve_targetFC";
                    Arcpy.CopyFeatures(fc_path, targetFC, true);

                    // 获取融合字段
                    string dissolvePath = $@"{gdb_path}\dissolve";
                    string sql = "";
                    foreach (var fieldName in fieldNames)
                    {
                        sql += fieldName + ";";
                    }
                    string new_sql = sql.Remove(sql.Length - 1);
                    // 融合
                    Arcpy.Dissolve(targetFC, dissolvePath, new_sql);
                    // 要素转线
                    string ex_line = $@"{gdb_path}\ex_line";
                    Arcpy.FeatureToLine(dissolvePath, ex_line);

                    string layer = "dissolve_targetFC";
                    string el_layer = layer;

                    // 按属性选择图层
                    Arcpy.SelectLayerByAttribute(layer, $"Shape_Area < {minArea}");
                    // 消除
                    Arcpy.Eliminate(el_layer, featureClass_path, "", ex_line, true);
                    // 移除图层
                    MapCtlTool.RemoveLayer(layer);

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
            string url = "https://blog.csdn.net/xcc34452366/article/details/137628819";
            UITool.Link2Web(url);
        }


        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_field);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_field);
        }


        private void combox_fc_DropClose(object sender, EventArgs e)
        {
            try
            {
                string fcPath = combox_fc.ComboxText();
                // 生成字段列表
                if (fcPath != "")
                {

                    UITool.AddFieldsToListBox(listbox_field, fcPath);
                }

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private List<string> CheckData(string featureClass_path)
        {
            List<string> result = new List<string>();

            string result_value = CheckTool.CheckGDBFeature(featureClass_path);
            if (result_value != "")
            {
                result.Add(result_value);
            }

            return result;
        }


        private void openFeatureClassButton_Click(object sender, RoutedEventArgs e)
        {
            textFeatureClassPath.Text = UITool.SaveDialogFeatureClass();
        }

    }
}
