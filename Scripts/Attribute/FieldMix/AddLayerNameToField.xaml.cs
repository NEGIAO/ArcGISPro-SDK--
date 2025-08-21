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

namespace CCTool.Scripts.Attribute.FieldMix
{
    /// <summary>
    /// Interaction logic for AddLayerNameToField.xaml
    /// </summary>
    public partial class AddLayerNameToField : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public AddLayerNameToField()
        {
            InitializeComponent();
            UITool.AddFeatureLayersAndTablesToListbox(listbox_fc);
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "添加图层名称和路径到字段";

        private async void btn_go_click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                bool isAddName = (bool)checkBox_name.IsChecked;
                bool isAddPath = (bool)checkBox_path.IsChecked;
                bool isAddFcName = (bool)checkBox_fcName.IsChecked;

                string fieldName = txt_name.Text;
                string fieldPath = txt_path.Text;
                string fieldFcName = txt_fcName.Text;

                // 文本空值处理
                if (fieldName == "") { fieldName = "图层名"; }
                if (fieldPath == "") { fieldPath = "路径"; }
                if (fieldFcName == "") { fieldPath = "要素名"; }


                // 判断参数是否选择完全
                if (isAddName == false && isAddPath ==false && isAddFcName == false)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }
                if (listbox_fc.Items.Count==0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 获取要素列表
                List<string> list_layer = listbox_fc.ItemsAsString();

                await QueuedTask.Run(() =>
                {
                    foreach (string layer in list_layer)
                    {
                        string layer_single = layer.GetLayerSingleName();
                        // 去除数字标记
                        if (layer_single.Contains('：'))
                        {
                            layer_single = layer_single[..layer_single.IndexOf("：")];
                        }

                        pw.AddMessageStart($"处理要素或表：{layer_single}");
                        // 添加图层名称
                        if (isAddName)
                        {
                            pw.AddMessageMiddle(5, $"添加字段：{fieldName}", Brushes.Gray);
                            // 添加字段
                            Arcpy.AddField(layer, fieldName, "TEXT");
                            // 计算字段
                            Arcpy.CalculateField(layer, fieldName, $"'{layer_single}'");
                        }
                        // 添加图层路径
                        if (isAddPath)
                        {
                            pw.AddMessageMiddle(5, $"添加字段：{fieldPath}", Brushes.Gray);
                            // 获取路径
                            string path = layer.TargetLayerPath().Replace(@"\",@"\\");
                            // 添加字段
                            Arcpy.AddField(layer, fieldPath, "TEXT");
                            // 计算字段
                            Arcpy.CalculateField(layer, fieldPath, $"'{path}'");
                        }
                        // 添加要素名称
                        if (isAddFcName)
                        {
                            pw.AddMessageMiddle(5, $"添加字段：{fieldFcName}", Brushes.Gray);
                            // 获取要素名
                            string fcName = layer.TargetFcName().Split('.')[0];
                            // 添加字段
                            Arcpy.AddField(layer, fieldFcName, "TEXT");
                            // 计算字段
                            Arcpy.CalculateField(layer, fieldFcName, $"'{fcName}'");
                        }
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

        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_fc);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_fc);
        }

        private void btn_help_click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135625991?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }
    }
}
