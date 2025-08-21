using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.UI.ProMapTool;
using Microsoft.Office.Core;
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
using Geometry = ArcGIS.Core.Geometry.Geometry;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for GroupBSM.xaml
    /// </summary>
    public partial class GroupBSM : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public GroupBSM()
        {
            InitializeComponent();

            // 读取参数
            textFolderPath.Text = BaseTool.ReadValueFromReg("GroupBSM", "folderPath");
            textFieldName.Text = BaseTool.ReadValueFromReg("GroupBSM", "fieldName");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "图斑聚类分组";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                var defGDB = Project.Current.DefaultGeodatabasePath;

                // 获取参数
                string folderPath = textFolderPath.Text;
                string fieldName = textFieldName.Text;

                // 判断参数是否选择完全
                if (folderPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 存储参数
                BaseTool.WriteValueToReg("GroupBSM", "folderPath", folderPath);
                BaseTool.WriteValueToReg("GroupBSM", "fieldName", fieldName);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 异步执行
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("获取所有shp文件");
                    // 获取所有shp文件
                    var shpFiles = DirTool.GetAllFiles(folderPath, ".shp");

                    string dissolve = $@"{defGDB}\dissolve";
                    foreach (string shpFile in shpFiles)
                    {
                        // shp名称
                        string shpName = shpFile.TargetFcName();
                        pw.AddMessageMiddle(100 / shpFiles.Count, $"处理__{shpName}");

                        // 添加1个标记字段
                        Arcpy.AddField(shpFile, fieldName, "LONG");
                        // 融合
                        Arcpy.Dissolve(shpFile, dissolve, "", "SINGLE_PART");

                        // 标记
                        // 获取原始图层和标识图层的要素类
                        FeatureClass originFeatureClass = shpFile.TargetFeatureClass();
                        FeatureClass identityFeatureClass = dissolve.TargetFeatureClass();

                        // 获取标记图层的要素游标
                        using RowCursor identityCursor = identityFeatureClass.Search();
                        // 遍历源图层的要素
                        while (identityCursor.MoveNext())
                        {
                            using Feature identityFeature = (Feature)identityCursor.Current;

                            // 标识图层的OID
                            long oid = long.Parse(identityFeature["OBJECTID"].ToString());

                            // 获取源要素的几何
                            Geometry identityGeometry = identityFeature.GetShape();

                            // 创建空间查询过滤器，以获取与源要素有重叠的目标要素
                            SpatialQueryFilter spatialFilter = new SpatialQueryFilter
                            {
                                FilterGeometry = identityGeometry,
                                SpatialRelationship = SpatialRelationship.Intersects
                            };

                            // 获取源图层的要素游标
                            using RowCursor originCursor = originFeatureClass.Search(spatialFilter);
                            while (originCursor.MoveNext())
                            {
                                using Feature originFeature = (Feature)originCursor.Current;
                                // 获取目标要素的几何
                                Geometry originGeometry = originFeature.GetShape();

                                // 赋值
                                originFeature[fieldName] = oid;

                                originFeature.Store();
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

        private void openSHPButton_Click(object sender, RoutedEventArgs e)
        {
            textFolderPath.Text = UITool.OpenDialogFolder();
        }


        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/145090617";
            UITool.Link2Web(url);
        }
    }
}
