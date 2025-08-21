using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.SS.Formula.Functions;
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

namespace CCTool.Scripts.UI.ProWindow
{
    /// <summary>
    /// Interaction logic for AddFields.xaml
    /// </summary>
    public partial class AddFields : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public AddFields()
        {
            InitializeComponent();
            //textFolderPath.Text = @"D:\【软件资料】\GIS相关\GisPro工具箱\4-测试数据\空 - 副本";
            //textExcelPath.Text = @"C:\Users\Administrator\Desktop\添加字段工具示例Excel.xlsx";
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "添加字段（批量）";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                var gdb = Project.Current.DefaultGeodatabasePath;
                // 获取工程默认文件夹位置
                var def_path = Project.Current.HomeFolderPath;

                List<string> shps = UITool.GetStringFromListbox(listbox_shp);
                // 获取指标
                string folder_path = textFolderPath.Text;
                string excel_path = textExcelPath.Text;

                // 判断参数是否选择完全
                if (folder_path == "" || excel_path == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 异步执行
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

                    pw.AddMessageMiddle(10, "获取字段属性结构表");

                    // 获取字段属性结构表
                    List<List<string>> list_field_attribute = new List<List<string>>();
                    // 获取名称、别名、类型、长度
                    List<string> list_mc = ExcelTool.GetListFromExcelAll(excel_path, 0, 2);
                    List<string> list_bm = ExcelTool.GetListFromExcelAll(excel_path, 1, 2);
                    List<string> list_fieldType = ExcelTool.GetListFromExcelAll(excel_path, 2, 2);
                    List<string> list_lenth = ExcelTool.GetListFromExcelAll(excel_path, 3, 2);
                    // 加入集合
                    for (int i = 0; i < list_mc.Count; i++)
                    {
                        list_field_attribute.Add(new List<string> { list_mc[i], list_bm[i], list_fieldType[i], list_lenth[i] });
                    }

                    pw.AddMessageMiddle(10, "获取所有要素类及表文件");

                    // 添加字段
                    foreach (string shp in shps)
                    {
                        string file = $@"{folder_path}{shp}";
                        pw.AddMessageMiddle(5, $"添加字段:{shp}");
                        // 如果不含关键字，直接添加字段
                        foreach (var fa in list_field_attribute)
                        {
                            Arcpy.AddField(file, fa[0], fa[2], fa[1], fa[3].ToInt());
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

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.OpenDialogExcel();
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textFolderPath.Text = UITool.OpenDialogFolder();

            try
            {
                // 更新listbox
                UpdateSHP();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        // 更新listbox
        public async void UpdateSHP()
        {
            string folder = textFolderPath.Text;

            // 清除listbox
            listbox_shp.Items.Clear();
            // 生成SHP要素列表
            if (textFolderPath.Text != "")
            {
                // 合并路径列表
                List<string> obj_all = new List<string>();
                // 获取所有shp文件
                List<string> shps =  DirTool.GetAllFiles(folder, ".shp");
                // 添加shp
                obj_all.AddRange(shps);

                // 获取所有gdb文件
                List<string> gdbs = DirTool.GetAllGDBFilePaths(folder);
                // 获取所有要素类文件
                if (gdbs is not null)
                {
                    foreach (var gdb in gdbs)
                    {
                        List<string> fcs_and_tbs = await QueuedTask.Run(() =>
                        {
                            return gdb.GetFeatureClassAndTablePath();
                        });
                        // 添加要素类和表
                        obj_all.AddRange(fcs_and_tbs);
                    }
                }

                foreach (var file in obj_all)
                {
                    // 将shp文件做成checkbox放入列表中
                    CheckBox cb = new CheckBox();
                    cb.Content = file.Replace(folder, "");
                    cb.IsChecked = true;
                    listbox_shp.Items.Add(cb);

                }
            }
        }



        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135666064?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        private void downLoadButton_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://pan.baidu.com/s/1hm9aIzbgO_J5kIv94udzNA?pwd=swh3";
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
            UITool.SelectListboxItems(listbox_shp);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_shp);
        }
    }
}
