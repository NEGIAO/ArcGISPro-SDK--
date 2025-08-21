using ApeFree.DataStore;
using ApeFree.DataStore.Core;
using ApeFree.DataStore.Local;
using ArcGIS.Desktop.Core;
using Aspose.Words.Lists;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Managers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
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

namespace CCTool.Scripts.DataPross.TXT
{
    /// <summary>
    /// Interaction logic for TitleMessage.xaml
    /// </summary>
    public partial class TitleMessage : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        string conName = "TitleBox";

        public TitleMessage()
        {
            InitializeComponent();

            // 初始化Title列表
            for (int i = 1; i < 7; i++)
            {
                listbox_title.Items.Add($"抬头文本_{i}");
            }
            listbox_title.SelectedIndex = 0;

            // 初始化title框
            string t1 = BaseTool.ReadValueFromReg(conName, "抬头文本_1");
            txtBox_head.Text = t1;
        }

        private void btn_read_Click(object sender, RoutedEventArgs e)
        {
            // 获取当前文本
            string text = txtBox_head.Text;
            string select = listbox_title.SelectedItems[0].ToString();
            // 保存
            BaseTool.WriteValueToReg(conName, "initTitle", text);
            BaseTool.WriteValueToReg(conName, select, text);

            // 更新主面板上的文本显示
            EventCenter.Broadcast(EventDefine.UpdataTitle);
            Close();
        }

        private void listbox_title_selectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 当前Title框的内容
            string title = txtBox_head.Text;

            // 变动前的选择项
            var lastSelected = e.RemovedItems;
            // 变动后的选择项
            var nextSelected = e.AddedItems;


            if (lastSelected != null && lastSelected.Count > 0)
            {
                string lastItem = lastSelected[0].ToString();
                // 将变动前的选项内容写入配置文件
                BaseTool.WriteValueToReg(conName, lastItem, title);
            }


            if (nextSelected != null && nextSelected.Count > 0)
            {
                string nextItem = nextSelected[0].ToString();
                // 读取变动后选项的配置属性到文本框
                string next = BaseTool.ReadValueFromReg(conName, nextItem);
                txtBox_head.Text = next;
            }
        }
    }
}
