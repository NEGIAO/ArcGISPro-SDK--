using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers;
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

namespace CCTool.Scripts.Attribute.FieldString
{
    /// <summary>
    /// Interaction logic for ZfillZero.xaml
    /// </summary>
    public partial class ZfillZero : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "ZfillZero";
        public ZfillZero()
        {
            InitializeComponent();

            textBox_len.Text = BaseTool.ReadValueFromReg(toolSet, "len").ToInt(6).ToString();
            rb_zfill.IsChecked = BaseTool.ReadValueFromReg(toolSet, "isZfill").ToBool("true");
            rb_clear.IsChecked = BaseTool.ReadValueFromReg(toolSet, "isClear").ToBool("false");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "文本前面补零去零";

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

                int len = textBox_len.Text.ToInt();

                bool isZfill = (bool)rb_zfill.IsChecked;
                bool isClear = (bool)rb_clear.IsChecked;

                // 判断参数是否选择完全
                if (layer_path == "" || fieldName == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }
                // 前缀保存到本地
                BaseTool.WriteValueToReg(toolSet, "len", len);
                BaseTool.WriteValueToReg(toolSet, "isZfill", isZfill);
                BaseTool.WriteValueToReg(toolSet, "isClear", isClear);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(layer_path, fieldName, len);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }


                    // 补齐零
                    if (isZfill)
                    {
                        pw.AddMessageMiddle(20, $"文本前面补零至{len}位");
                        Arcpy.CalculateField(layer_path, fieldName, $"!{fieldName}!.zfill({len})");
                    }
                    // 去零
                    else
                    {
                        pw.AddMessageMiddle(20, "文本前面移除多余的零");
                        Arcpy.CalculateField(layer_path, fieldName, $"int(!{fieldName}!)");
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
            string url = "https://blog.csdn.net/xcc34452366/article/details/145531867";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string in_data, string field, int len)
        {
            List<string> result = new List<string>();

            int fieldLength = in_data.GetFieldAtt(field).Length;
            // 检查len是否超过字段本身长度
            if (len > fieldLength)
            {
                result.Add($"输入的文本长度大于【{field}】字段的长度，改小一点！");
            }

            return result;
        }

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field);
        }


    }
}
