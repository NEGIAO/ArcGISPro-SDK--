using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
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
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CCTool.Scripts.MixApp.StyleMix
{
    /// <summary>
    /// Interaction logic for SortStylxItem.xaml
    /// </summary>
    public partial class SortStylxItem : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public SortStylxItem()
        {
            InitializeComponent();

            combox_type.Items.Add("面符号");
            combox_type.Items.Add("线符号");
            combox_type.Items.Add("点符号");
            combox_type.SelectedIndex = 0;
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

                List<string> newList = UITool.GetTextFromListBox(listbox_item);

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
                    // 获取SymbolStyleItem
                    List<SymbolStyleItem> allStyleItems = StylxTool.GetSymbolStyleItem(styleProjectItem);
                    // 先清空
                    foreach (SymbolStyleItem styleItem in allStyleItems)
                    {
                        StyleHelper.RemoveItem(styleProjectItem, styleItem);
                    }

                    // 排序
                    List<SymbolStyleItem> newStyleItems = new List<SymbolStyleItem>();

                    foreach (string name in newList)
                    {
                        foreach (SymbolStyleItem styleItem in allStyleItems)
                        {
                            if (styleItem.Name == name)
                            {
                                newStyleItems.Add(styleItem);
                            }
                        }
                    }


                    // 再按排序结果添加
                    foreach (SymbolStyleItem styleItem in newStyleItems)
                    {
                        StyleHelper.AddItem(styleProjectItem, styleItem);
                    }
                });
                MessageBox.Show($"样式库{stylxName}顺序更新完成!");
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147100224";
            UITool.Link2Web(url);
        }

        private void btn_bottom_Click(object sender, RoutedEventArgs e)
        {
            UITool.ListboxItemBottom(listbox_item);
        }

        private void btn_top_Click(object sender, RoutedEventArgs e)
        {
            UITool.ListboxItemTop(listbox_item);
        }

        private void btn_up_Click(object sender, RoutedEventArgs e)
        {
            UITool.ListboxItemUp(listbox_item);
        }

        private void btn_down_Click(object sender, RoutedEventArgs e)
        {
            UITool.ListboxItemDown(listbox_item);
        }

        private void combox_stylx_DropClosed(object sender, EventArgs e)
        {
            UpdataList();
        }


        private void combox_type_Closed(object sender, EventArgs e)
        {
            UpdataList();
        }

        private void btn_sort_Click(object sender, RoutedEventArgs e)
        {
            SortList();
        }

        // 更新列表
        private async void UpdataList()
        {
            try
            {
                // 样式库名称
                string stylxName = combox_stylx.ComboxText();
                // 样式类型
                string stylxType = combox_type.Text;

                // 如果条件没达到
                if (stylxName == "")
                {
                    return;
                }

                // 先清空
                listbox_item.Items.Clear();


                List<string> names = await QueuedTask.Run(() =>
                {
                    StyleItemType styleItemType = stylxType switch
                    {
                        "面符号" => StyleItemType.PolygonSymbol,
                        "线符号" => StyleItemType.LineSymbol,
                        "点符号" => StyleItemType.PointSymbol,
                        _ => StyleItemType.Unknown,
                    };


                    // 获取StyleProjectItem
                    StyleProjectItem styleProjectItem = stylxName.TargetStyleProjectItem();
                    // 获取名称
                    return StylxTool.GetStyleItemNames(styleProjectItem, styleItemType);
                });

                foreach (string name in names)
                {
                    listbox_item.Items.Add(name);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        // 按名称排序
        private async void SortList()
        {
            try
            {
                // 获取现在的列表
                List<string> names = await QueuedTask.Run(() =>
                {
                    return UITool.GetTextFromListBox(listbox_item);
                });

                // 重新排序
                names.Sort();

                // 先清空
                listbox_item.Items.Clear();

                foreach (string name in names)
                {
                    listbox_item.Items.Add(name);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }
    }
}
