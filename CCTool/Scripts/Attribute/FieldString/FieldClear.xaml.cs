using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using System;
using System.Collections.Generic;
using System.IO;
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

namespace CCTool.Scripts.UI.ProWindow
{
    /// <summary>
    /// Interaction logic for FieldClear.xaml
    /// </summary>
    public partial class FieldClear : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public FieldClear()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "清洗字段值";

        private void combox_fc_DropOpen(object sender, EventArgs e)
        {
            UITool.AddFeatureLayerAndTableToComboxPlus(combox_fc);
        }

        private async void btn_go_click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string input_fc = combox_fc.ComboxText();
                // 符号系统选择
                string model = "";
                if (rb_string_clearSpace.IsChecked == true) { model = "string_clearSpace"; }
                else if (rb_string_clearNone.IsChecked == true) { model = "string_clearNone"; }
                else if (rb_num_none2zero.IsChecked == true) { model = "num_none2zero"; }
                else if (rb_num_zero2none.IsChecked == true) { model = "num_zero2none"; }

                // 判断参数是否选择完全
                if (input_fc == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 获取参数listbox
                List<string> fieldNames = UITool.GetCheckboxStringFromListBox(listbox_field);

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart($"获取字段");

                    if (model == "string_clearSpace")        // 字符串：清除空格
                    {
                        foreach (var field in fieldNames)
                        {
                            pw.AddMessageMiddle(10, $"清除字符串空格__{field}", Brushes.Gray);
                            FieldCalTool.ClearTextSpace(input_fc, field);
                        }
                    }
                    else if (model == "string_clearNone")        // 字符串：清除空值
                    {
                        foreach (var field in fieldNames)
                        {
                            pw.AddMessageMiddle(10, $"清除字符串null值__{field}", Brushes.Gray);
                            FieldCalTool.ClearTextNull(input_fc, field);
                        }
                    }
                    else if (model == "num_none2zero")        // 数值：空值转0
                    {
                        foreach (var field in fieldNames)
                        {
                            pw.AddMessageMiddle(10, $"清除数字型null值__{field}", Brushes.Gray);
                            FieldCalTool.ClearMathNull(input_fc, field);
                        }
                    }
                    else if (model == "num_zero2none")        // 数值：0转空值
                    {
                        foreach (var field in fieldNames)
                        {
                            pw.AddMessageMiddle(10, $"0值转null值__{field}", Brushes.Gray);
                            FieldCalTool.Zero2Null(input_fc, field);
                        }
                    }
                    pw.AddMessageEnd();
                });
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_help_click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135753052?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        private void combox_fc_DropClosed(object sender, EventArgs e)
        {
            try
            {
                string fcPath = combox_fc.ComboxText();

                if (fcPath == "")
                {
                    return;
                }

                // 根据模式生成字段列表

                // 符号系统选择
                string model = "";
                if (rb_string_clearSpace.IsChecked == true) { model = "string_clearSpace"; }
                else if (rb_string_clearNone.IsChecked == true) { model = "string_clearNone"; }
                else if (rb_num_none2zero.IsChecked == true) { model = "num_none2zero"; }
                else if (rb_num_zero2none.IsChecked == true) { model = "num_zero2none"; }

                if (model == "string_clearSpace" || model == "string_clearNone")
                {
                    UITool.AddTextFieldsToListBox(listbox_field, fcPath);
                }
                else
                {
                    UITool.AddAllMathFieldsToListBox(listbox_field, fcPath);
                }


            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_field);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_field);
        }
    }
}
