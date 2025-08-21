using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
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
using static System.Net.Mime.MediaTypeNames;

namespace CCTool.Scripts.UI.ProWindow
{
    /// <summary>
    /// Interaction logic for UpdateAliasName.xaml
    /// </summary>
    public partial class UpdateAliasName : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public UpdateAliasName()
        {
            InitializeComponent();
            UITool.AddFeatureLayersToListbox(listbox_fc);
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "更新字段别名(属性映射)";

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            // 打开Excel文件
            string path = UITool.OpenDialogExcel();
            // 将Excel文件的路径置入【textExcelPath】
            textExcelPath.Text = path;
        }

        private async void btn_go_click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string excel_path = textExcelPath.Text;

                // 判断参数是否选择完全
                if (excel_path == "" || listbox_fc.Items.Count == 0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 获取要素列表
                List<string> list_fc = new List<string>();
                foreach (CheckBox item in listbox_fc.Items)
                {
                    if (item.IsChecked == true)
                    {
                        list_fc.Add(item.Content.ToString());
                    }
                }

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData();
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10,"获取对照表");
                    // 获取对照表
                    Dictionary<string, string> dic = ExcelTool.GetDictFromExcel(excel_path);

                    foreach (var fc in list_fc)
                    {
                        pw.AddMessageMiddle(2, $"处理要素或表：{fc}");
                        // 获取所选图层的所有字段
                        var fields = GisTool.GetFieldsFromTarget(fc);

                        // 更改字段别名
                        foreach (var field in fields)
                        {
                            string fieldName = field.Name;
                            if (dic.ContainsKey(fieldName))
                            {
                                pw.AddMessageMiddle(2, @$"更改字段别名：{fieldName}");
                                // 更改字段
                                Arcpy.AlterField(fc, fieldName, fieldName, dic[fieldName]);
                            }
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
            string url = "https://blog.csdn.net/xcc34452366/article/details/135665185?spm=1001.2014.3001.5502";
            UITool.Link2Web(url);
        }

        private void downLoadButton_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://pan.baidu.com/s/1kyocm7sV5Q1ytQ5hBsmK_A?pwd=383e";
            UITool.Link2Web(url);
        }

        private List<string> CheckData()
        {
            List<string> result = new List<string>();

            // 检查是否正常提取Excel
            string result_excel = CheckTool.CheckExcelPick();
            if (result_excel != "")
            {
                result.Add(result_excel);
            }

            return result;
        }

        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_fc);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_fc);
        }
    }
}
