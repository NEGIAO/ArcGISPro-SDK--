using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using ArcGIS.Desktop.Mapping;
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
using CCTool.Scripts.ToolManagers;
using NPOI.OpenXmlFormats.Vml;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Extensions;

namespace CCTool.Scripts.UI.ProWindow
{
    /// <summary>
    /// Interaction logic for MergerCAD.xaml
    /// </summary>
    public partial class MergerCAD : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public MergerCAD()
        {
            InitializeComponent();
            combox_type.Items.Add("点");
            combox_type.Items.Add("线");
            combox_type.Items.Add("面");
            combox_type.Items.Add("文字");
            combox_type.SelectedIndex = 1;
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "批量CAD合并为要素";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                var def_gdb = Project.Current.DefaultGeodatabasePath;
                // 获取指标
                string folder_path = textFolderPath.Text;
                string featureClass_path = textFeatureClassPath.Text;

                // 获取所有CAD文件
                List<string> files = UITool.GetStringFromListbox(listbox_shp);

                // 判断参数是否选择完全
                if (folder_path == "" || featureClass_path == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 获取输出要素类型
                var featureclass_type = combox_type.Text switch
                {
                    "点" => "Point",
                    "线" => "Polyline",
                    "面" => "Polygon",
                    "文字" => "Annotation",
                    _ => "",
                };

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 异步执行
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(featureClass_path);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10,"获取所有CAD文件");
                    
                    // 初始化一个输出要素列表
                    List<string> list_out_fc = new List<string>();
                    // 分解文件夹目录，获取文件名和路径字段值
                    int num = 1;
                    foreach (var file in files)
                    {
                        // 获取CAD文件名
                        string cad_name = file.Substring(file.LastIndexOf(@"\") + 1).Replace(".dwg", "");
                        pw.AddMessageMiddle(5, $"解析CAD文件：{cad_name}");

                        // 定义输出要素名称
                        string out_fc = $@"{def_gdb}\TransForm{num}_{featureclass_type}";
                        string target_fc = $@"{folder_path}{file}\{featureclass_type}";
                        // 复制要素
                        Arcpy.CopyFeatures(target_fc, out_fc);

                        // 加入列表
                        list_out_fc.Add(out_fc);
                        num++;
                    }
                    pw.AddMessageMiddle(10, "合并要素");
                    // 合并要素
                    Arcpy.Merge(list_out_fc, featureClass_path);
                    // 删除中间要素
                    foreach (var out_fc in list_out_fc)
                    {
                        Arcpy.Delect(out_fc);
                    }
                    // 将要素类添加到当前地图
                    MapCtlTool.AddLayerToMap(featureClass_path);

                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textFolderPath.Text = UITool.OpenDialogFolder();

            // 填写输出路径
            textFeatureClassPath.Text = Project.Current.DefaultGeodatabasePath + @"\CAD合并";

            try
            {
                // 更新listbox
                UpdateCAD();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        // 更新listbox
        public void UpdateCAD()
        {
            string folder_path = textFolderPath.Text;
            // 清除listbox
            listbox_shp.Items.Clear();
            // 生成SHP要素列表
            if (textFolderPath.Text != "")
            {
                // 获取所有CAD文件
                var files = DirTool.GetAllFiles(folder_path , ".dwg");
                foreach (var file in files)
                {
                    // 将shp文件做成checkbox放入列表中
                    CheckBox cb = new CheckBox();
                    cb.Content = file.Replace(folder_path, "");
                    cb.IsChecked = true;
                    listbox_shp.Items.Add(cb);

                }
            }
        }


        private void openFeatureClassButton_Click(object sender, RoutedEventArgs e)
        {
            textFeatureClassPath.Text = UITool.SaveDialogFeatureClass();
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135841060?spm=1001.2014.3001.5502";
            UITool.Link2Web(url);
        }

        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_shp);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_shp);
        }

        private List<string> CheckData(string featureClassPath)
        {
            List<string> result = new List<string>();

            // 检查gdb要素路径是否合理
            string result_value = CheckTool.CheckGDBPath(featureClassPath);
            if (result_value != "")
            {
                result.Add(result_value);
            }

            string result_value2 = CheckTool.CheckGDBIsNumeric(featureClassPath);
            if (result_value2 != "")
            {
                result.Add(result_value);
            }
            return result;
        }
    }
}
