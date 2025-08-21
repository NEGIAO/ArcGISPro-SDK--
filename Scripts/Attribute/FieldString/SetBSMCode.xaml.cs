using ArcGIS.Core.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
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

namespace CCTool.Scripts.Attribute.FieldString
{
    /// <summary>
    /// Interaction logic for SetBSMCode.xaml
    /// </summary>
    public partial class SetBSMCode : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "BSMTool";

        public SetBSMCode()
        {
            InitializeComponent();

            UITool.InitFieldToComboxPlus(combox_field, "BSM", "string");

            textBox_front.Text = BaseTool.ReadValueFromReg(toolSet, "front");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "BSM编码";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayerAndTableToComboxPlus(combox_fc);
        }

        private async void btn_go_click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取数据
                string layer_path = combox_fc.ComboxText();
                string fieldName = combox_field.ComboxText();

                string front = textBox_front.Text;
                bool isByLength = (bool)rb_bsmLength.IsChecked;
                bool isSort = (bool)checkBox_sort.IsChecked;

                string leng = textBox_len.Text ?? "";     // 自定义字段长度
                int BLength = int.Parse(leng);

                // 获取默认数据库
                var gdb = Project.Current.DefaultGeodatabasePath;

                // 判断参数是否选择完全
                if (layer_path == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }
                // 前缀保存到本地
                BaseTool.WriteValueToReg(toolSet, "front", front);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(layer_path, fieldName);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(20, "获取OID字段");
                    // 获取OID字段
                    string oidField = layer_path.TargetIDFieldName();

                    // 字段长度
                    int len = BLength;
                    if (isByLength)
                    {
                        len = GisTool.GetFieldFromString(layer_path, fieldName).Length;
                    }
                    // 是否要重排
                    if (isSort)
                    {
                        pw.AddMessageMiddle(20, "按空间位置重新排序【左上至右下】");
                        // 全选
                        MapCtlTool.SelectAllFeature(layer_path);
                        // 排序
                        Arcpy.Sort(layer_path, gdb + @"\sort_fc", "Shape ASCENDING", "UL");

                        // 获取图层及原要素的路径
                        FeatureLayer featureLayer = layer_path.TargetFeatureLayer();
                        string fc_path = layer_path.TargetLayerPath();
                        // 获取符号系统
                        CIMRenderer cr = featureLayer.GetRenderer();
                        // 更新要素
                        Arcpy.CopyFeatures(gdb + @"\sort_fc", fc_path, true);
                        // 应用符号系统
                        featureLayer.SetRenderer(cr);
                    }

                    pw.AddMessageMiddle(30, $"计算{fieldName}");
                    // 计算BSM，区分shp和gdb
                    string source = layer_path.TargetLayerPath();
                    if (source.Contains(".gdb"))
                    {
                        Arcpy.CalculateField(layer_path, fieldName, $"'{front}'+'0' * ({len} - len(str(!{oidField}!))-{front.Length}) + str(!{oidField}!)");
                    }
                    else
                    {
                        Arcpy.CalculateField(layer_path, fieldName, $"'{front}'+'0' * ({len} - len(str(!{oidField}!+1))-{front.Length}) + str(!{oidField}!+1)");
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

        private void btn_help_click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/136716272?spm=1001.2014.3001.5502";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string in_data, string field)
        {
            List<string> result = new List<string>();

            // 检查是否有BSM字段
            bool isHave = GisTool.IsHaveFieldInTarget(in_data, field);
            if (!isHave)
            {
                result.Add($"图层属性表不包含【{field}】字段！");
            }

            return result;
        }

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field);
        }
    }
}
