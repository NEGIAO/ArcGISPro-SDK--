using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Documents;
using System.Windows;
using CCTool.Scripts.LayerPross;
using CCTool.Scripts.Manager;

namespace CCTool.Scripts.ToolManagers.Extensions
{
    public static class UIExtension
    {
        /// <returns>指定的单元格</returns>  
        public static DataGridCell GetCell(this DataGrid dataGrid, int rowIndex, int columnIndex)
        {
            DataGridRow rowContainer = dataGrid.GetRow(rowIndex);
            if (rowContainer != null)
            {
                DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(rowContainer);
                DataGridCell cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex);
                if (cell == null)
                {
                    dataGrid.ScrollIntoView(rowContainer, dataGrid.Columns[columnIndex]);
                    cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex);
                }
                return cell;
            }
            return null;
        }

        /// <returns>指定的行号</returns>  
        public static DataGridRow GetRow(this DataGrid dataGrid, int rowIndex)
        {
            DataGridRow rowContainer = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex);
            if (rowContainer == null)
            {
                dataGrid.UpdateLayout();
                dataGrid.ScrollIntoView(dataGrid.Items[rowIndex]);
                rowContainer = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex);
            }
            return rowContainer;
        }

        /// <returns>第一个指定类型的子可视对象</returns>  
        public static T GetVisualChild<T>(Visual parent) where T : Visual
        {
            T child = default(T);
            int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < numVisuals; i++)
            {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent, i);
                child = v as T;
                if (child == null)
                {
                    child = GetVisualChild<T>(v);
                }
                if (child != null)
                {
                    break;
                }
            }
            return child;
        }

        // 获取RichTextBox的文本内容
        public static string GetRichText(this RichTextBox richTextBox)
        {
            // 创建一个 TextRange 对象
            TextRange textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);

            // 获取 RichTextBox 的纯文本内容
            string text = textRange.Text;

            return text;
        }

        // 富文本添加信息框文字
        public static void AddMessage(this RichTextBox tb_message, string add_text, SolidColorBrush solidColorBrush = null)
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
            });
        }

        // 从ListBox【选中的CheckBox】中获取string列表
        public static List<string> ItemsAsString(this ListBox listBox)
        {
            List<string> strList = new List<string>();

            foreach (CheckBox item in listBox.Items)
            {
                if (item.IsChecked == true)
                {
                    strList.Add(item.Content.ToString());
                }
            }
            return strList;
        }

        // 获取加强型combox的文本
        public static string ComboxText(this ComboBox combox)
        {
            // 获取参数
            ComboBoxContent flc_fc = (ComboBoxContent)combox.SelectedItem;
            string fc_path = "";
            if (flc_fc is not null) { fc_path = flc_fc.Name; }

            return fc_path;
        }


        // 获取加强型listbox的文本
        public static List<string> ListboxText(this ListBox listbox)
        {
            // 获取参数
            List<string> names = new List<string>();
            foreach (var item in listbox.Items)
            {
                ListBoxContent flc_fc = (ListBoxContent)item;

                if (flc_fc.IsSelect == true) { names.Add(flc_fc.Name); }
            }

            return names;
        }

        // 获取ListBox中的文字列表
        public static List<string> GetCheckListBoxText(this ListBox lb)
        {
            List<string> list = new List<string>();

            // 获取所有选中的txt
            foreach (CheckBox item in lb.Items)
            {
                if (item.IsChecked == true)
                {
                    list.Add(item.Content.ToString());
                }
            }

            return list;
        }
    }
}
