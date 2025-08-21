using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
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
    /// Interaction logic for ClipGDB.xaml
    /// </summary>
    public partial class ClipGDB : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public ClipGDB()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "按范围分割数据库";

        private void openResultGDBButton_Click(object sender, RoutedEventArgs e)
        {
            textResultGDB.Text = UITool.OpenDialogFolder();
        }

        private void openOriginalGDBButton_Click(object sender, RoutedEventArgs e)
        {
            textOriginalGDB.Text = UITool.OpenDialogGDB();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string original_gdb = textOriginalGDB.Text;
                string clip_fc = comboxClipFeature.ComboxText();
                string clip_field = comboxClipField.ComboxText();
                string folder_resultGDB = textResultGDB.Text;

                // 工程默认文件夹位置
                string folder_path = Project.Current.HomeFolderPath;

                // 判断参数是否选择完全
                if (original_gdb == "" || clip_fc == "" || clip_field == "" || folder_resultGDB == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(clip_fc, clip_field);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10, "解析范围要素");
                    // 创建临时数据库
                    Arcpy.CreateFileGDB(folder_path, "新数据库");
                    string new_gdb = $@"{folder_path}\新数据库.gdb";

                    // 分解范围要素
                    Arcpy.SplitByAttributes(clip_fc, new_gdb, clip_field);
                    // 收集分割后的地块
                    List<string> list_area = new_gdb.GetFeatureClassPathFromGDB();
                    // 分区处理
                    foreach (var area in list_area)
                    {
                        // 获取area_name
                        string area_name = area[(area.LastIndexOf(@"\")+1)..];

                        pw.AddMessageMiddle(10, $"分割范围:{area_name}");
                        // 目标gdb路径
                        string gdbPath = folder_resultGDB + @$"\{area_name}.gdb";

                        // 复制GDB
                        DirTool.CopyAllFiles(original_gdb, gdbPath);
                        // 裁剪要素类
                        List<string> list_fc = original_gdb.GetFeatureClassPathFromGDB();
                        foreach (var fc in list_fc)
                        {
                            pw.AddMessageMiddle(2, $"{fc[(fc.LastIndexOf(@"\") + 1)..]}", Brushes.Gray);

                            // 目标要素路径
                            string targetPath = fc.Replace(original_gdb, gdbPath);
                            // 裁剪
                            Arcpy.Clip(fc, area, targetPath);

                            // 更改别名
                            string aliasName = fc.TargetFeatureClass().GetDefinition().GetAliasName();
                            string fcName = targetPath[(targetPath.LastIndexOf(@"\") + 1)..];
                            GisTool.AlterAliasName(gdbPath, fcName, aliasName);
                        }
                        // 裁剪栅格
                        List<string>list_raster = original_gdb.GetRasterPath();
                        foreach (var raster in list_raster)
                        {
                            pw.AddMessageMiddle(2, $"{raster[(raster.LastIndexOf(@"\") + 1)..]}", Brushes.Gray);
                            Arcpy.RasterClip(raster, raster.Replace(original_gdb, folder_resultGDB + @$"\{area_name}.gdb"), area);
                        }
                    }

                    // 删除中间数据
                    Directory.Delete(new_gdb, true);
                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void comboxClipFeature_DropOpen(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(comboxClipFeature);
        }

        private void comboxClipField_DropOpen(object sender, EventArgs e)
        {
            string clip = comboxClipFeature.ComboxText();
            UITool.AddTextFieldsToComboxPlus(clip, comboxClipField);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135816584?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string clip_fc, string clip_field)
        {
            List<string> result = new List<string>();
            // 检查是否包含空值
            string result_value_cs = CheckTool.CheckFieldValueEmpty(clip_fc, clip_field);
            if (result_value_cs != "")
            {
                result.Add(result_value_cs);
            }

            return result;
        }
    }
}
