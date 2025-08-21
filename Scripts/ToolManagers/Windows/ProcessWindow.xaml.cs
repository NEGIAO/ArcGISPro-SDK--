using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static System.Net.Mime.MediaTypeNames;
using ArcGIS.Desktop.Framework.Dialogs;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;
using SixLabors.ImageSharp.ColorSpaces;

namespace CCTool.Scripts.Manager
{
    /// <summary>
    /// Interaction logic for ProcessWindow.xaml
    /// </summary>
    public partial class ProcessWindow : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 初始时间
        readonly DateTime startTime = DateTime.Now;

        public ProcessWindow()
        {
            InitializeComponent();
        }

        // 变更进度条的进度【100%】
        private void AddProcess(double percent)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                pb.Value += percent;
            });
        }

        // 添加信息框文字_时间
        private void AddTime(DateTime time_base)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                DateTime time_now = DateTime.Now;
                TimeSpan time_span = time_now - startTime;
                string time_total = time_span.ToString()[..time_span.ToString().LastIndexOf(".")];
                string add_text = "………………用时" + time_total + "\r";

                // 创建一个新的TextRange对象，范围为新添加的文字
                TextRange newRange = new TextRange(tb_message.Document.ContentEnd, tb_message.Document.ContentEnd)
                {
                    Text = add_text
                };
                // 设置新添加文字的颜色为灰色
                newRange.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.Gray);
                // 设置新添加文字的样式为斜体
                newRange.ApplyPropertyValue(TextElement.FontStyleProperty, FontStyles.Italic);
            });
        }

        // 添加信息框文字_时间_最终
        private void AddTimeEnd(SolidColorBrush solidColorBrush = null)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                DateTime time_now = DateTime.Now;
                TimeSpan time_span = time_now - startTime;
                string time_total = time_span.ToString()[..time_span.ToString().LastIndexOf(".")];
                string add_text = "\r总用时：" + time_total;

                // 创建一个新的TextRange对象，范围为新添加的文字
                TextRange newRange = new TextRange(tb_message.Document.ContentEnd, tb_message.Document.ContentEnd)
                {
                    Text = add_text
                };
                // 设置新添加文字的颜色为灰色
                newRange.ApplyPropertyValue(TextElement.ForegroundProperty, solidColorBrush);
            });
        }

        // 添加信息框文字
        private void AddMessage(string add_text, SolidColorBrush solidColorBrush = null)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
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

                // 如果有结束语句,lb显示工具执行结束
                if (add_text.Contains("工具执行完成"))
                {
                    lb.Content = "工具执行完成！";
                }
            });
        }


        ///  外部调用
        // 综合显示进度【抬头】
        public void AddMessageTitle(string tool_name)
        {
            AddMessage($"开始执行【{tool_name}】工具…………{startTime}", Brushes.Green);
        }

        // 综合显示进度【起始】
        public void AddMessageStart(string add_text, SolidColorBrush solidColorBrush = null)
        {
            AddMessage("\r" + add_text, solidColorBrush);
        }

        // 综合显示进度【中间】
        public void AddMessageMiddle(double percent, string add_text, SolidColorBrush solidColorBrush = null)
        {
            AddProcess(percent);
            AddTime(startTime);
            AddMessage(add_text, solidColorBrush);
        }

        // 综合显示进度【结束】
        public void AddMessageEnd()
        {
            AddProcess(100);
            AddTime(startTime);
            AddMessage("工具执行完成！！！", Brushes.Blue);
            // 总用时
            AddTimeEnd(Brushes.Blue);
        }

    }
}
