using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using ArcGIS.Desktop.Mapping.Events;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using SharpCompress.Common;
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

namespace CCTool.Scripts.CusTool2
{
    /// <summary>
    /// Interaction logic for Dat2FeatureClass.xaml
    /// </summary>
    public partial class Dat2FeatureClass : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public Dat2FeatureClass()
        {
            InitializeComponent();

            Init();
        }

        // 初始化
        private void Init()
        {
            // combox_sr框中添加几种预制坐标系
            combox_sr.Items.Add("WGS_1984");
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

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "DAT文件转要素类(批量)";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string folder_path = textFolderPath.Text;
                string fc_path = textFeatureClassPath.Text;
                string spatial_reference = combox_sr.Text;


                var cb_txts = listbox_txt.Items;

                // 判断参数是否选择完全
                if (folder_path == "" || fc_path == "" || spatial_reference == "" || cb_txts.Count == 0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 获取目标数据库和点要素名
                string gdbPath = fc_path[..(fc_path.IndexOf(".gdb") + 4)];
                string fcName = fc_path[(fc_path.LastIndexOf(@"\") + 1)..];

                // 判断要素名是不是以数字开头
                bool isNum = fcName.IsNumeric();
                if (isNum)
                {
                    MessageBox.Show("输出的要素名不规范，不能以数字开头！");
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
                    // 检查数据
                    List<string> errs = CheckData(fc_path);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    // 参数获取
                    string gdb_def = Project.Current.DefaultGeodatabasePath;
                    pw.AddMessageMiddle(10, " 创建一个空要素");

                    // 创建一个空要素
                    Arcpy.CreateFeatureclass(gdb_def, "tem_fc", "POINT", spatial_reference);
                    string targetFC = gdb_def + @"\tem_fc";

                    pw.AddMessageMiddle(10, "新建字段");
                    // 新建字段
                    GisTool.AddField(targetFC, "文件名");
                    GisTool.AddField(targetFC, "文件路径");
                    GisTool.AddField(targetFC, "Z坐标", FieldType.Double);

                    pw.AddMessageMiddle(10, " 获取所有txt文件");

                    // 打开数据库
                    using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_def))))
                    {
                        // 创建要素并添加到要素类中
                        using FeatureClass featureClass = gdb.OpenDataset<FeatureClass>("tem_fc");
                        // 解析txt文件内容，创建面要素
                        foreach (string path in list_txtPath)
                        {
                            string shp_name = path[(path.LastIndexOf(@"\") + 1)..];
                            pw.AddMessageMiddle(2, " 处理：" + shp_name);

                            // 读取文件内容
                            string[] texts = File.ReadAllLines(path);

                            // 一个文件可能有多要素
                            foreach (string txt in texts)
                            {
                                // 排除无效行
                                if (!txt.Contains(','))
                                {
                                    continue;
                                }

                                string[] contents = txt.Split(',');

                                // XYZ坐标
                                double X = double.Parse(contents[2]);
                                double Y = double.Parse(contents[3]);
                                double Z = double.Parse(contents[4]);

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
                                    rowBuffer["文件名"] = shp_name.Replace(".dat", "");
                                    rowBuffer["文件路径"] = path;
                                    rowBuffer["Z坐标"] = Z;

                                    MapPointBuilderEx pb = new MapPointBuilderEx(X, Y);

                                    // 给新添加的行设置形状
                                    rowBuffer[featureClassDefinition.GetShapeField()] = pb.ToGeometry();

                                    // 在表中创建新行
                                    using Feature feature = featureClass.CreateRow(rowBuffer);
                                    context.Invalidate(feature);      // 标记行为无效状态
                                }, featureClass);

                                // 执行编辑操作
                                editOperation.Execute();
                            }
                        }
                    }
                    // 保存编辑
                    Project.Current.SaveEditsAsync();

                    // 复制要素
                    Arcpy.CopyFeatures(gdb_def + @"\tem_fc", fc_path);

                    // 将要素类添加到当前地图
                    MapCtlTool.AddLayerToMap(fc_path);

                    pw.AddMessageEnd();
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
            textFeatureClassPath.Text = UITool.SaveDialogFeatureClass();
        }

        private void openSHPButton_Click(object sender, RoutedEventArgs e)
        {

            // 打开TXT文件夹
            string folder = UITool.OpenDialogFolder();
            textFolderPath.Text = folder;
            // 更新列表
            UpadataListBox();

            // 填写输出要素路径
            textFeatureClassPath.Text = Project.Current.DefaultGeodatabasePath + @"\DAT文件转要素类";
        }

        private void UpadataListBox()
        {
            string pick = txt_pick.Text;    // 通配符
            string folder = textFolderPath.Text;
            // 清除listbox
            listbox_txt.Items.Clear();
            // 生成TXT要素列表
            if (folder != "")
            {
                // 获取所有shp文件
                List<string> newList = new List<string>();
                List<string> files = DirTool.GetAllFiles(folder, ".dat");
                // 通配符，不包含就移除
                if (pick != "")
                {
                    foreach (string file in files)
                    {
                        if (file.Contains(pick))
                        {
                            newList.Add(file);
                        }
                    }
                }
                else
                {
                    newList = files;
                }

                // 将txt文件做成checkbox放入列表中
                foreach (var file in newList)
                {
                    CheckBox cb = new CheckBox();
                    cb.Content = file.Replace(folder, "");
                    cb.IsChecked = true;
                    listbox_txt.Items.Add(cb);
                }
            }
        }

        private List<string> CheckData(string featurePath)
        {
            List<string> result = new List<string>();

            // 判断输出路径是否为gdb
            string gdbRusult = CheckTool.CheckGDBPath(featurePath);
            if (gdbRusult != "")
            {
                result.Add(gdbRusult);
            }

            return result;
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/141637345?spm=1001.2014.3001.5502";
            UITool.Link2Web(url);
        }

        private void txt_pick_Changed(object sender, TextChangedEventArgs e)
        {
            // 更新列表
            UpadataListBox();
        }
    }
}