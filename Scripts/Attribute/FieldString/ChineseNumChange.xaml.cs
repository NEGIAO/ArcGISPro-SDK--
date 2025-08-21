using ArcGIS.Core.Data;
using ArcGIS.Desktop.Framework.Threading.Tasks;
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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CCTool.Scripts.Attribute.FieldString
{
    /// <summary>
    /// Interaction logic for ChineseNumChange.xaml
    /// </summary>
    public partial class ChineseNumChange : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public ChineseNumChange()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "中文数字与阿拉伯数字互转";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayerAndTableToComboxPlus(combox_fc);
        }

        private async void btn_go_click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取数据
                string in_data = combox_fc.ComboxText();
                string input_field = combox_field_input.ComboxText();
                string output_field = combox_field_output.ComboxText();

                bool chineseToNum = (bool)rb_cn_num.IsChecked;

                // 判断参数是否选择完全
                if (in_data == "" || input_field == "" || output_field == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 转换模式，提示
                if (chineseToNum) { pw.AddMessageMiddle(20, $"中文数字转阿拉拍数字"); }
                else { pw.AddMessageMiddle(20, "阿拉拍数字转中文数字"); }

                await QueuedTask.Run(() =>
                {
                    Table tb = in_data.TargetTable();
                    using RowCursor rowCursor = tb.Search();
                    while (rowCursor.MoveNext())
                    {
                        Row row = rowCursor.Current;
                        // 获取字段值
                        var inputField = row[input_field];

                        if (inputField is not null)
                        {
                            string in_value = inputField.ToString();
                            // 转换
                            // 中文数字转阿拉拍数字
                            if (chineseToNum)
                            {
                                string result_value = BaseTool.ChineseConverToNum(in_value);
                                row[output_field] = result_value;
                            }
                            // 阿拉拍数字转中文数字
                            else
                            {
                                string result_value = BaseTool.NumConverToChinese(in_value);
                                row[output_field] = result_value;
                            }

                            row.Store();
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

        private void btn_help_click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/145533567";
            UITool.Link2Web(url);
        }

        private void combox_field_input_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field_input);
        }

        private void combox_field_output_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field_output);
        }

    }
}
