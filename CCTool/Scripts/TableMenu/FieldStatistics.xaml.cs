using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Dialogs;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

namespace CCTool.Scripts.TableMenu
{
    /// <summary>
    /// Interaction logic for FieldStatistics.xaml
    /// </summary>
    public partial class FieldStatistics : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public FieldStatistics()
        {
            InitializeComponent();

            // 将选择的字段填入字段框
            var tableView = TableView.Active;
            string selectedField = tableView.GetSelectedFields().FirstOrDefault();

            UITool.InitFieldToComboxPlus(combox_field, selectedField, "float");

            // 统计字段
            StatisticsArea();
        }

        // 统计字段
        private async void StatisticsArea()
        {
            // 接收参数
            string fieldName = combox_field.ComboxText();

            if (fieldName == "")
            {
                return;
            }

            try
            {
                string strAll = "";
                await QueuedTask.Run(() =>
                {
                    // 获取当前激活的表格视图
                    var tableView = TableView.Active;
                    if (tableView == null || tableView.MapMember == null)
                    {
                        return;
                    }

                    // 获取表格
                    RowCursor rowCursor;
                    if (tableView.MapMember is FeatureLayer featureLayer)
                    {
                        rowCursor = featureLayer.TargetSelectCursor();
                    }
                    else if (tableView.MapMember is StandaloneTable standaloneTable)
                    {
                        rowCursor = standaloneTable.TargetSelectCursor();
                    }
                    else
                    {
                        rowCursor = null;
                    }

                    // 统计值
                    double total = 0;
                    double min = 9999999999999999999;
                    double max = -8999999999999999999;
                    long count = 0;

                    // 统计
                    while (rowCursor.MoveNext())
                    {
                        Row row = rowCursor.Current;

                        var value = row[fieldName];
                        if (value is not null)
                        {
                            // 提取数字，计入统计
                            double va = GetNumber(value);

                            // 总和
                            total += va;
                            // 最大值
                            if (va>max)
                            {
                                max = va;
                            }
                            // 最小值
                            if (va < min)
                            {
                                min = va;
                            }
                        }
                        // 计数加1
                        count++;

                    }

                    // 文字汇合
                    strAll += $"总和：{total}\n最大值：{max}\n最小值：{min}\n平均值：{total/count}\n计数：{count}\n";

                });

                // 写入文本框
                text_statistics.Text = strAll;
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        private void combox_field_DropClosed(object sender, EventArgs e)
        {
            // 统计字段
            StatisticsArea();
        }

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddAllTableMathFieldsToComboxPlus(combox_field);
        }

        private void btn_help_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147431405?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        private void copyButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            string txt = text_statistics.Text;
            string re = txt[(txt.IndexOf("总和：")+3)..txt.IndexOf("\n")];
            // 复制到剪贴板
            Clipboard.SetText(re);
        }

        private void refreshButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // 统计字段
            StatisticsArea();
        }

        // 获取数字
        private double GetNumber(object str)
        {
            double result = 0;

            // 如果是科学计数法的数字
            if (str.ToString().Contains("E-"))
            {
                result = (double)str;
            }
            else
            {
                result = str.ToString().GetWord("小数").ToDouble();
            }

            return result;
        }
    }
}
