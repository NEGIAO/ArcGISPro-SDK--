using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Dml.Diagram;
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

namespace CCTool.Scripts.MixApp.StyleMix
{
    /// <summary>
    /// Interaction logic for ExchangeStylxValue.xaml
    /// </summary>
    public partial class ExchangeStylxValue : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public ExchangeStylxValue()
        {
            InitializeComponent();
        }

        private void combox_stylx_DropDown(object sender, EventArgs e)
        {
            UITool.AddStylxsToComboxPlus(combox_stylx);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string stylxName = combox_stylx.ComboxText();

                // 判断参数是否选择完全
                if (stylxName == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                Close();
                await QueuedTask.Run(() =>
                {

                    // 获取StyleProjectItem
                    StyleProjectItem styleProjectItem = stylxName.TargetStyleProjectItem();
                    // 替换
                    StylxTool.ChangeValueAndTag(styleProjectItem);

                });
                
                MessageBox.Show($"样式库{stylxName}的名称和标签对换完成!");
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147044747";
            UITool.Link2Web(url);
        }
    }
}
