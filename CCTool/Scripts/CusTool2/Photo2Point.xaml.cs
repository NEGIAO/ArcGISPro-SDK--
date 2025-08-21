using ActiproSoftware.Windows.Extensions;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Data.Exceptions;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.GeoProcessing;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells.Drawing;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Shared;
using NPOI.OpenXmlFormats.Vml;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
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
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;
using static System.Net.Mime.MediaTypeNames;
using CheckBox = System.Windows.Controls.CheckBox;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

namespace CCTool.Scripts.CusTool2
{
    /// <summary>
    /// Interaction logic for Photo2Point.xaml
    /// </summary>
    public partial class Photo2Point : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public Photo2Point()
        {
            InitializeComponent();
            Init();
        }

        // 初始化
        private void Init()
        {
            // combox_sr框中添加几种预制坐标系
            combox_sr.Items.Add("GCS_WGS_1984");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_25");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_26");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_27");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_28");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_29");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_30");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_31");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_32");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_33");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_34");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_35");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_36");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_37");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_38");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_39");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_40");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_41");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_42");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_43");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_44");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_45");

            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_75E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_78E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_81E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_84E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_87E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_90E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_93E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_96E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_99E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_102E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_105E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_108E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_111E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_114E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_117E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_120E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_123E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_126E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_129E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_132E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_135E");

            combox_sr.SelectedIndex = 0;

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "照片TXT文件转SHP点";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string folder_path = textFolderPath.Text;
                string def_gdb = textFeatureClassPath.Text;
                string spatial_reference = combox_sr.Text;

                var cb_txts = listbox_txt.Items;

                // 判断参数是否选择完全
                if (folder_path == "" || def_gdb == "" || spatial_reference == "" || cb_txts.Count == 0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 获取所有选中的txt
                List<string> list_txtPath = new List<string>();
                foreach (CheckBox shp in cb_txts)
                {
                    if (shp.IsChecked == true)
                    {
                        list_txtPath.Add(folder_path + shp.Content);
                    }
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                // 异步执行
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");

                    foreach (string txtPath in list_txtPath)
                    {
                        string shp_name_old = txtPath[(txtPath.LastIndexOf(@"\") + 1)..].Replace(".txt", "");  // 获取要素名
                        // 预处理一下要素名，避免一些奇奇怪怪的符号
                        string shp_name = shp_name_old.Replace(".", "_");

                        // 要素名不能以数字开头
                        bool isNum = shp_name.IsNumeric();
                        if (isNum)
                        {
                            shp_name = "T" + shp_name;
                        }

                        pw.AddMessageMiddle(10, $@"创建点要素：{shp_name}");

                        // 创建一个空要素
                        Arcpy.CreateFeatureclass(def_gdb, shp_name, "POINT", spatial_reference);
                        // 新建字段
                        Arcpy.AddField(@$"{def_gdb}\{shp_name}.shp", "照片名称", "TEXT");
                        Arcpy.AddField(@$"{def_gdb}\{shp_name}.shp", "纬度", "Double");
                        Arcpy.AddField(@$"{def_gdb}\{shp_name}.shp", "经度", "Double");
                        Arcpy.AddField(@$"{def_gdb}\{shp_name}.shp", "Yaw", "Double");
                        Arcpy.AddField(@$"{def_gdb}\{shp_name}.shp", "Pitch", "Double");
                        Arcpy.AddField(@$"{def_gdb}\{shp_name}.shp", "Roll", "Double");
                        Arcpy.AddField(@$"{def_gdb}\{shp_name}.shp", "水平精度", "Double");
                        Arcpy.AddField(@$"{def_gdb}\{shp_name}.shp", "垂直精度", "Double");

                        // 打开gdb
                        using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(def_gdb)));
                        using FeatureClass featureClass = gdb.OpenDataset<FeatureClass>(shp_name);
   
                        // 获取txt文件的文本内容
                        string text = TxtTool.GetTXTContent(txtPath);

                        // 获取坐标点文本
                        string fcs_text = text.Split("垂直精度\r\n")[1];
                        // 去除第一部分非坐标文本
                        string[] textList = fcs_text.Split("\r\n");

                        foreach (string tt in textList)
                        {
                            if (!tt.Contains(',')) { continue; }

                            string[] PtAtt = tt.Split(",");

                            /// 构建面要素
                            // 创建编辑操作对象
                            EditOperation editOperation = new EditOperation();
                            editOperation.Callback(context =>
                            {
                                // 获取要素定义
                                FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();
                                // 创建RowBuffer
                                using RowBuffer rowBuffer = featureClass.CreateRowBuffer();

                                // 写入字段值
                                rowBuffer["照片名称"] = PtAtt[0];
                                rowBuffer["纬度"] = PtAtt[1];
                                rowBuffer["经度"] = PtAtt[2];
                                rowBuffer["Yaw"] = PtAtt[3];
                                rowBuffer["Pitch"] = PtAtt[4];
                                rowBuffer["Roll"] = PtAtt[5];
                                rowBuffer["水平精度"] = PtAtt[7];
                                rowBuffer["垂直精度"] = PtAtt[8];

                                MapPointBuilderEx pb = new(double.Parse(PtAtt[2]), double.Parse(PtAtt[1]));

                                // 给新添加的行设置形状
                                rowBuffer[featureClassDefinition.GetShapeField()] = pb.ToGeometry();

                                // 在表中创建新行
                                using Feature feature = featureClass.CreateRow(rowBuffer);
                                context.Invalidate(feature);      // 标记行为无效状态
                            }, featureClass);

                            // 执行编辑操作
                            editOperation.Execute();

                        }

                        // 保存编辑
                        Project.Current.SaveEditsAsync();
                    }

                    pw.AddMessageMiddle(10, "工具运行完成！！！");
                });
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        private void openFeatureClassButton_Click(object sender, RoutedEventArgs e)
        {
            textFeatureClassPath.Text = UITool.OpenDialogGDB();
        }

        private void openSHPButton_Click(object sender, RoutedEventArgs e)
        {
            // 打开TXT文件夹
            string folder = UITool.OpenDialogFolder();
            textFolderPath.Text = folder;
            // 清除listbox
            listbox_txt.Items.Clear();
            // 生成TXT要素列表
            if (textFolderPath.Text != "")
            {
                // 获取所有shp文件
                var files = DirTool.GetAllFiles(folder, ".txt");
                foreach (var file in files)
                {
                    // 将txt文件做成checkbox放入列表中
                    CheckBox cb = new CheckBox();
                    cb.Content = file.Replace(folder, "");
                    cb.IsChecked = true;
                    listbox_txt.Items.Add(cb);
                }
            }

        }


    }
}
