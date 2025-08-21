using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
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
using Table = ArcGIS.Core.Data.Table;

namespace CCTool.Scripts.DataPross.GDB
{
    /// <summary>
    /// Interaction logic for MergeGDB.xaml
    /// </summary>
    public partial class MergeGDB : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "MergeGDB";
        public MergeGDB()
        {
            InitializeComponent();

            textFolderPath.Text = BaseTool.ReadValueFromReg(toolSet, "gdbFolder");
            text_gdbName.Text = BaseTool.ReadValueFromReg(toolSet, "gdbName");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "合并GDB数据库";

        private void openForldrButton_Click(object sender, RoutedEventArgs e)
        {
            textFolderPath.Text = UITool.OpenDialogFolder();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string gdbFolder = textFolderPath.Text;
                string gdbName = text_gdbName.Text;

                // 判断参数是否选择完全
                if (gdbFolder == "" || gdbName == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 参数写入本地
                BaseTool.WriteValueToReg(toolSet, "gdbFolder", gdbFolder);
                BaseTool.WriteValueToReg(toolSet, "gdbName", gdbName);


                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("获取所有GDB文件");
                    // 获取所有GDB文件
                    List<string> gdbFiles = DirTool.GetAllGDBFilePaths(gdbFolder);
                    pw.AddMessageMiddle(10, "创建目标GDB");
                    // 创建合并GDB
                    string gdbPath = Arcpy.CreateFileGDB(gdbFolder, gdbName);
                    // 要素数据集列表
                    List<string> dataBaseNames = new List<string>();
                    // 要素类列表
                    List<string> featureClassNames = new List<string>();
                    // 独立表列表
                    List<string> tableNames = new List<string>();


                    foreach (string gdbFile in gdbFiles)
                    {
                        pw.AddMessageMiddle(10, $"处理数据库：{gdbFile}");
                        // 获取FeatureClass
                        using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbFile)));
                        // 获取要素数据集
                        IReadOnlyList<FeatureDatasetDefinition> featureDatases = gdb.GetDefinitions<FeatureDatasetDefinition>();
                        // 新建要素数据集
                        if (featureDatases.Count > 0)
                        {
                            foreach (var featureDatase in featureDatases)
                            {
                                string dbName = featureDatase.GetName();
                                if (!dataBaseNames.Contains(dbName))   // 如果是新的，就创建
                                {
                                    SpatialReference sr = featureDatase.GetSpatialReference();    // 数据集的坐标系
                                    double XYtolerance = sr.XYTolerance;        // 获取xy容差
                                    double XYResolution = sr.XYResolution;        // 获取xy分辨率
                                    double Ztolerance = sr.ZTolerance;        // 获取Z容差
                                    double Mtolerance = sr.MTolerance;        // 获取M容差

                                    Arcpy.CreateFeatureDataset(gdbPath, dbName, sr, XYtolerance, XYResolution, Ztolerance, Mtolerance);

                                    dataBaseNames.Add(dbName);
                                }

                            }
                        }

                        // 获取要素类
                        IReadOnlyList<FeatureClassDefinition> featureClasses = gdb.GetDefinitions<FeatureClassDefinition>();
                        if (featureClasses.Count > 0)
                        {
                            foreach (var featureClass in featureClasses)
                            {
                                string fcName = featureClass.GetName();
                                FeatureClass fc = gdb.OpenDataset<FeatureClass>(fcName);
                                // 获取要素类路径
                                string fcPath = fc.GetPath().ToString().Replace("file:///", "").Replace("/", @"\");
 
                                // 获取目标路径
                                string targetPath = gdbPath + fcPath[(fcPath.IndexOf(".gdb") + 4)..];

                                if (!featureClassNames.Contains(fcName))   // 如果是新的，就复制要素类
                                {
                                    Arcpy.CopyFeatures(fcPath, targetPath);

                                    // 复制后的要素别名丢失，修改回来
                                    string aliasName = fc.GetDefinition().GetAliasName();
                                    GisTool.AlterAliasName(gdbPath, fcName, aliasName);

                                    featureClassNames.Add(fcName);
                                }
                                else   // 如果已经有要素了，就追加
                                {
                                    Arcpy.Append(fcPath, targetPath);
                                }
                            }
                        }

                        // 获取独立表
                        IReadOnlyList<TableDefinition> tables = gdb.GetDefinitions<TableDefinition>();
                        // 新建独立表
                        if (tables.Count > 0)
                        {
                            foreach (var table in tables)
                            {
                                string tbName = table.GetName();
                                Table tb = gdb.OpenDataset<Table>(tbName);
                                // 获取独立表路径
                                string tbPath = tb.GetPath().ToString().Replace("file:///", "").Replace("/", @"\");
                                // 获取目标路径
                                string targetPath = gdbPath + tbPath[(tbPath.IndexOf(".gdb") + 4)..];

                                if (!tableNames.Contains(tbName))   // 如果是新的，就复制独立表
                                {
                                    Arcpy.CopyRows(tbPath, targetPath);

                                    // 复制后的要素别名丢失，修改回来
                                    string aliasName = tb.GetDefinition().GetAliasName();
                                    GisTool.AlterTableAliasName(gdbPath, tbName, aliasName);

                                    tableNames.Add(tbName);
                                }
                                else   // 如果已经有独立表了，就追加
                                {
                                    Arcpy.Append(tbPath, targetPath);
                                }
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

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135813877?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

    }
}
