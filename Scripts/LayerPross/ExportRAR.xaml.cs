using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Dml.Diagram;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
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

namespace CCTool.Scripts.LayerPross
{
    /// <summary>
    /// Interaction logic for ExportRAR.xaml
    /// </summary>
    public partial class ExportRAR : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public ExportRAR()
        {
            InitializeComponent();

            try
            {
                // 初始化rarName
                Layer layers = MapView.Active.GetSelectedLayers().FirstOrDefault();
                textRarName.Text = layers.Name;

                // 初始化文本框
                textFolderPath.Text = BaseTool.ReadValueFromReg("RARset", "path");
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
                return;
            }
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "图层导出压缩包";


        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string folderPath = textFolderPath.Text;
                string rarName = textRarName.Text;

                bool isGDB = (bool)rb_gdb.IsChecked;
                // 判断参数是否选择完全
                if (folderPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 获取当前选择的图层
                var layers = MapView.Active.GetSelectedLayers();
                // 获取当前选择的图层
                var tables = MapView.Active.GetSelectedStandaloneTables();

                if (layers == null && tables == null)
                {
                    MessageBox.Show("错误！请选择一个要素图层或表格！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 异步执行
                await QueuedTask.Run(() =>
                {
                    //  如是GDB
                    if (isGDB)
                    {
                        // 新建一个gdb
                        pw.AddMessageStart($"创建GDB数据库_{rarName}");
                        Arcpy.CreateFileGDB(folderPath, rarName);
                        string gdbPath = $@"{folderPath}\{rarName}.gdb";
                        // 导入要素图层
                        if (layers != null)
                        {
                            foreach (Layer layer in layers)
                            {
                                pw.AddMessageMiddle(10, $"导出要素_{layer.Name}", Brushes.Gray);
                                Arcpy.CopyFeatures(layer, $@"{gdbPath}\{layer.Name}");
                            }
                        }
                        // 导入独立表
                        if (tables != null)
                        {
                            foreach (StandaloneTable table in tables)
                            {
                                pw.AddMessageMiddle(10, $"导出独立表_{table.Name}", Brushes.Gray);
                                Arcpy.CopyRows(table, $@"{gdbPath}\{table.Name}");
                            }
                        }

                        pw.AddMessageMiddle(10, $"压缩文件");
                        // 压缩
                        ZipFile.CreateFromDirectory(gdbPath, $"{gdbPath}.Zip");
                        // 删除数据库
                        Directory.Delete(gdbPath, true);
                    }
                    else    //  如是SHP
                    {
                        // 新建一个文件夹
                        pw.AddMessageStart($"创建SHP文件夹_{rarName}");
                        string shpPath = @$"{folderPath}\{rarName}";
                        Directory.CreateDirectory(shpPath);
                        
                        // 导入要素图层
                        if (layers != null)
                        {
                            foreach (Layer layer in layers)
                            {
                                pw.AddMessageMiddle(10, $"导出要素_{layer.Name}.shp", Brushes.Gray);
                                Arcpy.CopyFeatures(layer, $@"{shpPath}\{layer.Name}.shp");
                            }
                        }

                        // 导入独立表
                        if (tables != null)
                        {
                            pw.AddMessageMiddle(10, $"独立表无法导出shp", Brushes.Red);
                        }

                        pw.AddMessageMiddle(10, $"压缩文件");
                        // 压缩
                        ZipFile.CreateFromDirectory(shpPath, $"{shpPath}.Zip");
                        // 删除数据库
                        Directory.Delete(shpPath, true);
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


        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/139088318";
            UITool.Link2Web(url);
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textFolderPath.Text = UITool.OpenDialogFolder();

            // 保存
            BaseTool.WriteValueToReg("RARset", "path", textFolderPath.Text);
        }
    }
}
