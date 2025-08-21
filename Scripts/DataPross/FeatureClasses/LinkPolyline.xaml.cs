using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
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

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    /// <summary>
    /// Interaction logic for LinkPolyline.xaml
    /// </summary>
    public partial class LinkPolyline : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "LinkPolyline";

        public LinkPolyline()
        {
            InitializeComponent();

            // 如果刚开始注册表没有值，就赋一个默认值
            text_distance.Text = BaseTool.ReadValueFromReg(toolSet, "minDistance", "0.01");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "断线连接";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc, "Polyline");
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string lyName = combox_fc.ComboxText();
                double minDistance = text_distance.Text.ToDouble();

                // 默认数据库位置
                var gdb_path = Project.Current.DefaultGeodatabasePath;
                // 工程默认文件夹位置
                string folder_path = Project.Current.HomeFolderPath;

                // 判断参数是否选择完全
                if (lyName == "" || minDistance == 0)
                {
                    MessageBox.Show("有必选参数为空，或参数填写错误！！！");
                    return;
                }

                BaseTool.WriteValueToReg(toolSet, "minDistance", minDistance);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageMiddle(10, "复制一个要素");
                    // 复制一个新的图层【去掉Z值】
                    string newLayer = @$"{gdb_path}\merge_{lyName}";
                    string fcName = $"merge_{lyName}";
                    Arcpy.CopyFeatures(lyName, newLayer, true, "Disabled");

                    FeatureLayer lineLayer = fcName.TargetFeatureLayer();

                    pw.AddMessageMiddle(20, "线段之间就近相连");

                    // 提取所有端点并记录所属线段ID
                    var endpoints = new List<(MapPoint Point, long ParentOID)>();

                    // 获取线图层中的所有线要素
                    using var featureCursor = lineLayer.Search();
                    var polylines = new List<(Polyline polyline, long OID)>();
                    while (featureCursor.MoveNext())
                    {
                        Feature feature = featureCursor.Current as Feature;
                        Polyline polyline = feature.GetShape() as Polyline;
                        polylines.Add((polyline, feature.GetObjectID()));
                    }

                    foreach (var polyline in polylines)
                    {
                        long oid = polyline.OID;

                        endpoints.Add((polyline.polyline.Points.First(), oid));
                        endpoints.Add((polyline.polyline.Points.Last(), oid));
                    }

                    // 3. 查找
                    List<List<long>> mergeLineIDs = new List<List<long>>();
                    List<long> allLineIDs = new List<long>();

                    foreach (var pointAtt in endpoints)
                    {
                        MapPoint point = pointAtt.Point;
                        long parentOID = pointAtt.ParentOID;

                        foreach (var targetPointAtt in endpoints)
                        {
                            MapPoint targetPoint = targetPointAtt.Point;
                            long targetParentOID = targetPointAtt.ParentOID;

                            if (targetParentOID == parentOID)
                            {
                                continue;
                            }
                            else
                            {
                                double distance = Math.Sqrt(Math.Pow(targetPoint.X - point.X, 2) + Math.Pow(targetPoint.Y - point.Y, 2));
                               
                                if (distance <= minDistance)    // 如果距离小于容差
                                {
                                    // 加入列表
                                    AddtoList(mergeLineIDs, allLineIDs, parentOID, targetParentOID);
                                }
                            }

                        }
                    }

                    // 合并
                    GisTool.MergeFeatures(lineLayer, mergeLineIDs);


                    // 保存编辑 
                    Project.Current.SaveEditsAsync();

                    // 刷新地图视图
                    MapView.Active.ZoomInFixed();

                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void AddtoList(List<List<long>> mergeLineIDs, List<long> allLineIDs, long parentOID, long targetParentOID)
        {
            // 如果都没有，就都加入
            if (!allLineIDs.Contains(parentOID) && !allLineIDs.Contains(targetParentOID))
            {
                mergeLineIDs.Add(new List<long>() { parentOID, targetParentOID });
                allLineIDs.Add(parentOID);
                allLineIDs.Add(targetParentOID);
            }
            // 如果有一个有，就加入另一个
            else if (allLineIDs.Contains(parentOID) && !allLineIDs.Contains(targetParentOID))
            {
                foreach (var mergeLineID in mergeLineIDs)
                {
                    if (mergeLineID.Contains(parentOID))
                    {
                        mergeLineID.Add(targetParentOID);
                        allLineIDs.Add(targetParentOID);
                    }
                }
            }

            else if (!allLineIDs.Contains(parentOID) && allLineIDs.Contains(targetParentOID))
            {
                foreach (var mergeLineID in mergeLineIDs)
                {
                    if (mergeLineID.Contains(targetParentOID))
                    {
                        mergeLineID.Add(parentOID);
                        allLineIDs.Add(parentOID);
                    }
                }
            }


        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/148613953";
            UITool.Link2Web(url);
        }
    }
}
