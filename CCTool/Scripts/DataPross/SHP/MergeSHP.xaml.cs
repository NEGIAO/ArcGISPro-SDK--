using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Vml;
using System;
using System.Collections.Generic;
using System.IO;
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
    /// Interaction logic for MergeSHP.xaml
    /// </summary>
    public partial class MergeSHP : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public MergeSHP()
        {
            InitializeComponent();

            combox_geoType.Items.Add("点");
            combox_geoType.Items.Add("线");
            combox_geoType.Items.Add("面");
            combox_geoType.SelectedIndex = 2;

        }
        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "合并文件夹下的所有SHP文件";

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
                string featureClass_path = textFeatureClassPath.Text;
                string shp_path = def_path + @"\ShpFiles";

                // 判断参数是否选择完全
                if (folder_path == "" || featureClass_path == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 获取目标数据库和点要素名
                string gdbPath = featureClass_path[..(featureClass_path.IndexOf(".gdb") + 4)];
                string fcName = featureClass_path[(featureClass_path.LastIndexOf(@"\") + 1)..];

                // 判断要素名是不是以数字开头
                bool isNum = fcName.IsNumeric();
                if (isNum)
                {
                    MessageBox.Show("输出的要素名不规范，不能以数字开头！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 异步执行
                await QueuedTask.Run(() =>
                {
                    // 复制shp文件夹用作后续处理
                    if (Directory.Exists(shp_path))
                    {
                        Directory.Delete(shp_path, true);
                    }
                    Directory.CreateDirectory(shp_path);

                    // 获取路径标签
                    string tab = folder_path.Substring(folder_path.LastIndexOf(@"\") + 1);
                    pw.AddMessageStart("解析shp文件，添加标记字段");
                    // 分解文件夹目录，获取文件名和路径字段值
                    foreach (string shp in shps)
                    {
                        // 获取名称和路径
                        string oldFile = folder_path + shp;
                        string shpFile= oldFile[(oldFile.LastIndexOf(@"\") + 1)..];
                        string shpName = shpFile[..shpFile.LastIndexOf(@".shp")];
                        string file = $@"{shp_path}\{shpFile}";
                        string path = oldFile.Replace(shpFile, "").Replace(@"\", @"/");

                        pw.AddMessageMiddle(0, shpFile);
                        // 复制一个
                        Arcpy.Copy(oldFile, file);

                        // 添加2个标记字段
                        Arcpy.AddField(file, "shpName", "TEXT");
                        Arcpy.AddField(file, "shpPath", "TEXT");
                        Arcpy.CalculateField(file, "shpName", "\"" + shpName + "\"");
                        Arcpy.CalculateField(file, "shpPath", "\"" + path + "\"");
                    }
                    pw.AddMessageMiddle(40, "合并要素");
                    // 合并要素
                    List<string> newFiles = DirTool.GetAllFiles(shp_path, ".shp");
                    Arcpy.Merge(newFiles, featureClass_path);
                    // 删除复制的shp文件夹
                    Directory.Delete(shp_path, true);
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

        private void openFeatureClassButton_Click(object sender, RoutedEventArgs e)
        {
            textFeatureClassPath.Text = UITool.SaveDialogFeatureClass();
        }

        private void openSHPButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                textFolderPath.Text = UITool.OpenDialogFolder();
                // 填写输出路径
                textFeatureClassPath.Text = Project.Current.DefaultGeodatabasePath + @"\SHP合并";

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
            // 类型
            string tp = combox_geoType.Text;
            GeometryType geoTp = tp switch
            {
                "点"=> GeometryType.Point,
                "线" => GeometryType.Polyline,
                "面" => GeometryType.Polygon,
                _ => GeometryType.Polygon,
            };

            string folder = textFolderPath.Text;

            // 清除listbox
            listbox_shp.Items.Clear();
            // 生成SHP要素列表
            if (textFolderPath.Text != "")
            {
                // 获取所有shp文件
                var files = DirTool.GetAllFiles(folder, ".shp");
                foreach (var file in files)
                {
                    GeometryType geoType = await QueuedTask.Run(() =>
                    {
                        return file.TargetGeoType();
                    });
                    // 如果符合类型要求
                    if (geoType == geoTp)
                    {
                        // 将shp文件做成checkbox放入列表中
                        CheckBox cb = new CheckBox();
                        cb.Content = file.Replace(folder, "");
                        cb.IsChecked = true;
                        listbox_shp.Items.Add(cb);
                    }
                    
                }
            }
        }


        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135623839?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        private void combox_geoType_Closed(object sender, EventArgs e)
        {
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
