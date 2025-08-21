using ArcGIS.Desktop.Core;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

namespace CCTool.Scripts.Attribute.FieldMix
{
    /// <summary>
    /// Interaction logic for CalculateFieldPlus.xaml
    /// </summary>
    public partial class CalculateFieldPlus : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public CalculateFieldPlus()
        {
            InitializeComponent();

            // 加载模式
            foreach (string model in models)
            {
                combox_model.Items.Add(model);
            }
            combox_model.SelectedIndex = 0;
        }
        // 计算模式
        public List<string> models = new List<string>()
        {
            "字段串切片_按起始终止index",
            "字段串切片_按起始终止文本",
            "度分秒转十进制度",
            //"十进制度转度分秒",
        };

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayerAndTableToComboxPlus(combox_fc);
        }

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddFieldsToComboxPlus(combox_fc.ComboxText(), combox_field);
        }

        private void t1_combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), t1_combox_field);
        }


        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string calField = combox_field.ComboxText();

                // 判断参数是否选择完全
                if (fc_path == "" || calField == "")
                {
                    MessageBox.Show("有必选参数为空或输入错误！！！");
                    return;
                }
                // 获取富文本的内容
                string exp = rich_text01.GetRichText();
                string code = rich_text02.GetRichText();

                await QueuedTask.Run(() =>
                {
                    // 计算字段
                    Arcpy.CalculateField(fc_path, calField, exp, code);
                });

                MessageBox.Show("计算字段完成！");

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/142145524";
            UITool.Link2Web(url);
        }

        private void combox_model_Closed(object sender, EventArgs e)
        {
            string model = combox_model.Text;
            foreach (TabItem tabItem in tc.Items)
            {
                if (tabItem.Header.ToString() == model)
                {
                    tc.SelectedItem = tabItem;
                }
            }
        }


        // 字段串切片_按起始终止index
        public async void StringClip1()
        {
            // 获取参数
            string fc_path = combox_fc.ComboxText();
            string calField = combox_field.ComboxText();

            // 判断参数是否选择完全
            if (fc_path == "" || calField == "")
            {
                MessageBox.Show("有必选参数为空或输入错误！！！");
                return;
            }
            // 获取富文本的内容
            string exp = rich_text01.GetRichText();

            await QueuedTask.Run(() =>
            {
                // 计算字段
                Arcpy.CalculateField(fc_path, calField, exp);
            });

            MessageBox.Show("计算字段完成！");
        }

        // 字段串切片_按起始终止文本
        public async void StringClip2()
        {
            // 获取参数
            string fc_path = combox_fc.ComboxText();
            string calField = combox_field.ComboxText();

            // 判断参数是否选择完全
            if (fc_path == "" || calField == "")
            {
                MessageBox.Show("有必选参数为空或输入错误！！！");
                return;
            }
            // 获取富文本的内容
            string exp = rich_text01.GetRichText();

            await QueuedTask.Run(() =>
            {
                // 计算字段
                Arcpy.CalculateField(fc_path, calField, exp);
            });

        }


        // 更新表达式【字段串切片_按起始终止index】
        public void UpdataFrame_StringClip1()
        {
            // 清空RichTextBox
            rich_text01.Document.Blocks.Clear();
            rich_text02.Document.Blocks.Clear();
            // 获取参数
            string t1_field = t1_combox_field.ComboxText();

            _ = int.TryParse(t1_textbox_start.Text, out int startIndex);
            _ = int.TryParse(t1_textbox_end.Text, out int endIndex);

            // 计算起终位置
            string start = "";
            string end = "";

            if (startIndex > 1)
            {
                start = (startIndex - 1).ToString();
            }

            if (endIndex > 1)
            {
                end = endIndex.ToString();
            }

            string exp = $"!{t1_field}![{start}:{end}]";

            // 更新文本
            AddMessage(rich_text01, exp);
            AddMessage(rich_text02, "");
        }

        // 更新表达式【字段串切片_按起始终止文本】
        public void UpdataFrame_StringClip2()
        {
            // 清空RichTextBox
            rich_text01.Document.Blocks.Clear();
            rich_text02.Document.Blocks.Clear();
            // 获取参数

            string t2_field = t2_combox_field.ComboxText();

            string startText = t2_textbox_start.Text;
            string endText = t2_textbox_end.Text;
            bool hasStart = (bool)cb_startText.IsChecked;
            bool hasEnd = (bool)cb_endText.IsChecked;

            // 计算表达式
            string startIndex = "";
            string endIndex = "";
            if (startText != "")
            {
                if (hasStart)
                {
                    startIndex = $"!{t2_field}!.find(\"{startText}\")";
                }
                else
                {
                    startIndex = $"!{t2_field}!.find(\"{startText}\")+len(\"{startText}\")";
                }
            }
            if (endText != "")
            {
                if (hasEnd)
                {
                    endIndex = $"!{t2_field}!.find(\"{endText}\")+len(\"{endText}\")";
                }
                else
                {
                    endIndex = $"!{t2_field}!.find(\"{endText}\")";
                }
            }

            string exp = $"!{t2_field}![{startIndex} : {endIndex}]";

            // 更新文本
            AddMessage(rich_text01, exp);
            AddMessage(rich_text02, "");
        }

        // 更新表达式【度分秒转十进制度】
        public void UpdataFrame_Degree1()
        {
            // 清空RichTextBox
            rich_text01.Document.Blocks.Clear();
            rich_text02.Document.Blocks.Clear();
            // 获取参数

            string t3_field = t3_combox_field.ComboxText();

            // 计算表达式
            string exp = $"ss(!{t3_field}!)";
            List<string> codes = new List<string>()
            {
                "def ss(in_text):",
                "    index1 = in_text.find(u'°')",
                "    index2 = in_text.find(u'′')",
                "    index3 = in_text.find(u'″')",
                "    degree = float(in_text[0:index1])",
                "    minutes = float(in_text[index1 + 1:index2])",
                "    seconds = float(in_text[index2 + 1:index3])",
                "    result = degree + minutes / 60 + seconds / 3600",
                "    return result",
            }; 

            // 更新文本
            AddMessage(rich_text01, exp);
            foreach (var code in codes)
            {
                AddMessage(rich_text02, code + "\r");
            }
        }

        // 更新表达式【十进制度转度分秒】
        public void UpdataFrame_Degree2()
        {
            // 清空RichTextBox
            rich_text01.Document.Blocks.Clear();
            rich_text02.Document.Blocks.Clear();
            // 获取参数

            string t4_field = t4_combox_field.ComboxText();

            // 计算表达式
            string exp = $"SS(!{t4_field}!)";
            string code = @"def ss(in_text):
    index1 = in_text.find(u'°')
    index2 = in_text.find(u'′')
    index3 = in_text.find(u'″')
    degree = float(in_text[0:index1])
    minutes = float(in_text[index1 + 1:index2])
    seconds = float(in_text[index2 + 1:index3])
    result = degree + minutes / 60 + seconds / 3600
    return result";

            // 更新文本
            AddMessage(rich_text01, exp);
            AddMessage(rich_text02, code);
        }

        // 添加信息框文字
        public void AddMessage(RichTextBox tb_message, string add_text, SolidColorBrush solidColorBrush = null)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                if (solidColorBrush == null)
                {
                    solidColorBrush = Brushes.Black;
                }
                // 创建一个新的TextRange对象，范围为新添加的文字
                TextRange newRange = new TextRange(tb_message.Document.ContentEnd, tb_message.Document.ContentEnd)
                {
                    Text = add_text
                };
                // 设置新添加文字的颜色
                newRange.ApplyPropertyValue(TextElement.ForegroundProperty, solidColorBrush);
                // 设置新添加文字的样式
                newRange.ApplyPropertyValue(TextElement.FontStyleProperty, FontStyles.Normal);

                //tb_message.Focus();        // RichTextBox获取焦点，有时也可以不用
                tb_message.ScrollToEnd();        // RichTextBox滚动到光标位置
            });
        }

        private void t1_combox_field_Closed(object sender, EventArgs e)
        {
            UpdataFrame_StringClip1();
        }

        private void t1_textbox_start_Change(object sender, TextChangedEventArgs e)
        {
            UpdataFrame_StringClip1();
        }

        private void t1_textbox_end_Change(object sender, TextChangedEventArgs e)
        {
            UpdataFrame_StringClip1();
        }

        private void t2_combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), t2_combox_field);
        }

        private void t2_combox_field_Closed(object sender, EventArgs e)
        {
            UpdataFrame_StringClip2();
        }

        private void t2_textbox_start_Change(object sender, TextChangedEventArgs e)
        {
            UpdataFrame_StringClip2();
        }

        private void t2_textbox_end_Change(object sender, TextChangedEventArgs e)
        {
            UpdataFrame_StringClip2();
        }

        private void cb_startText_Checked(object sender, RoutedEventArgs e)
        {
            UpdataFrame_StringClip2();
        }

        private void cb_endText_Checked(object sender, RoutedEventArgs e)
        {
            UpdataFrame_StringClip2();
        }

        private void t3_combox_field_Closed(object sender, EventArgs e)
        {
            UpdataFrame_Degree1();
        }

        private void t3_combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), t3_combox_field);
        }

        private void t4_combox_field_Closed(object sender, EventArgs e)
        {
            UpdataFrame_Degree2();
        }

        private void t4_combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), t4_combox_field);
        }
    }
}
