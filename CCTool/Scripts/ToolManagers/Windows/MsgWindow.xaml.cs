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

namespace CCTool.Scripts.ToolManagers.Windows
{
    /// <summary>
    /// Interaction logic for MsgWindow.xaml
    /// </summary>
    public partial class MsgWindow : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public MsgWindow()
        {
            InitializeComponent();
        }


        // 添加信息框文字
        public void AddMessage(string add_text, SolidColorBrush solidColorBrush = null)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                solidColorBrush ??= Brushes.Black;
                // 创建一个新的TextRange对象，范围为新添加的文字
                TextRange newRange = new(tb_message.Document.ContentEnd, tb_message.Document.ContentEnd)
                {
                    Text = add_text+"\r"
                };
                // 设置新添加文字的颜色
                newRange.ApplyPropertyValue(TextElement.ForegroundProperty, solidColorBrush);
                // 设置新添加文字的样式
                newRange.ApplyPropertyValue(TextElement.FontStyleProperty, FontStyles.Normal);
                // RichTextBox滚动到光标位置
                tb_message.ScrollToEnd();
            });
        }
    }
}
