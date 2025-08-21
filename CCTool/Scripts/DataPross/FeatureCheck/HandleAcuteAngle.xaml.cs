using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.GeoProcessing;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Dml.Diagram;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
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
using Geometry = ArcGIS.Core.Geometry.Geometry;

namespace CCTool.Scripts.LayerPross
{
    /// <summary>
    /// Interaction logic for HandleAcuteAngle.xaml
    /// </summary>
    public partial class HandleAcuteAngle : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "HandleAcuteAngle";

        public HandleAcuteAngle()
        {
            InitializeComponent();

            // 初始化参数选项
            txt_angle.Text = BaseTool.ReadValueFromReg(toolSet, "txt_angle");
            txt_distance.Text = BaseTool.ReadValueFromReg(toolSet, "txt_distance");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "自动处理尖锐角";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string layerName = combox_fc.ComboxText();

                double miniAngle = txt_angle.Text.ToDouble();
                double moveDistance = txt_distance.Text.ToDouble();

                string field_group = combox_field_group.ComboxText();

                string gdbPath = Project.Current.DefaultGeodatabasePath;

                // 判断参数是否正确填写
                if (miniAngle == 0 || moveDistance == 0)
                {
                    MessageBox.Show("有必选参数为空或填写错误！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "txt_angle", miniAngle);
                BaseTool.WriteValueToReg(toolSet, "txt_distance", moveDistance);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {

                    pw.AddMessageStart("复制图层");
                    // 获取图层
                    FeatureLayer featurelayer = layerName.TargetFeatureLayer();
                    // 复制目标图层
                    string targetPath = GisTool.CopyCosFeatureLayer(featurelayer);

                    // 获取gdbPath,targetName
                    string gdbPath = targetPath.TargetWorkSpace();
                    string targetName = targetPath.TargetFcName();

                    // 获取图层
                    FeatureLayer targetLayer = MapView.Active.Map.GetLayersAsFlattenedList().OfType<FeatureLayer>().FirstOrDefault(ly => ly.Name == targetName) as FeatureLayer;
                    // 获取目标的坐标系
                    SpatialReference sr = featurelayer.GetSpatialReference();

                    // OID字段
                    string oidField = targetLayer.TargetIDFieldName();

                    // 添加标记字段，并等于OID
                    string bjField = "处理标记";
                    GisTool.AddField(targetPath, bjField);
                    Arcpy.CalculateField(targetPath, bjField, $"!{oidField}!");

                    pw.AddMessageMiddle(20, "查找尖锐角，并分割");

                    // 获取图层的要素类
                    FeatureClass featureClass = featurelayer.GetFeatureClass();

                    // 遍历面要素类中的所有要素
                    RowCursor rowCursor = featurelayer.Search();

                    // 遍历图层的要素
                    while (rowCursor.MoveNext())
                    {
                        using Feature feature = (Feature)rowCursor.Current;

                        // oid字段值
                        long oid = long.Parse(feature[oidField].ToString());

                        // 获取要素的几何
                        Polygon polygon = feature.GetShape() as Polygon;
                        // 检查要素的类型
                        if (polygon is null)
                        {
                            MessageBox.Show("所选图层不是面要素！");
                            return;
                        }

                        // 检查角度
                        // 获取几何的所有段
                        var parts = polygon.Parts.ToList();

                        foreach (ReadOnlySegmentCollection collection in parts)
                        {
                            // 初始化点集
                            List<MapPoint> points = new List<MapPoint>();
                            // 每个环进行处理（第一个为外环，其它为内环）
                            foreach (Segment segment in collection)
                            {
                                MapPoint mapPoint = segment.StartPoint;     // 获取点
                                points.Add(mapPoint);
                            }
                            // 补充首末点，方便下一步计算
                            points.Add(points[0]);
                            points.Insert(0, points[points.Count - 2]);

                            // 计算角度
                            for (int i = 1; i < points.Count - 1; i++)
                            {
                                // 当前点及前后两个点
                                MapPoint p1 = points[i - 1];
                                MapPoint p2 = points[i];
                                MapPoint p3 = points[i + 1];

                                // 获取三点形成的夹角
                                double angle = GetAngle(p1, p2, p3);

                                // 符合条件的
                                if (angle < miniAngle)
                                {
                                    // p1到p2, p3到p2的距离
                                    double d_A = Math.Sqrt(Math.Pow((p1.X - p2.X), 2) + Math.Pow((p1.Y - p2.Y), 2));
                                    double d_B = Math.Sqrt(Math.Pow((p3.X - p2.X), 2) + Math.Pow((p3.Y - p2.Y), 2));
                                    // 如果折点间的距离存在小于设定好的移动距离的情况，就按最小距离
                                    double d_min = Math.Min(Math.Min(d_B, d_A), moveDistance);

                                    // 找出分割点
                                    MapPoint p1_cut = CalPosition(p2, p1, d_min, sr);
                                    MapPoint p2_cut = CalPosition(p2, p3, d_min, sr);

                                    var points_cut = new List<MapPoint>() { p1_cut, p2_cut };

                                    // 用分割点创建分割线
                                    Polyline cutLine = PolylineBuilderEx.CreatePolyline(points_cut, sr);

                                    // 创建编辑器
                                    var cutFeatures = new EditOperation();
                                    cutFeatures.Name = "Cut Features";

                                    // 执行（要素类通过oid选择出要素）
                                    cutFeatures.Split(targetLayer, oid, cutLine);

                                    if (!cutFeatures.IsEmpty)
                                    {
                                        var result = cutFeatures.Execute();
                                    }

                                    // 合并
                                    MergeFeature(targetName, bjField, oid, miniAngle, field_group);

                                }
                            }
                        }

                    }

                    // 取消选择
                    MapCtlTool.UnSelectAllFeature(targetName);
                    // 最终保存
                    Project.Current.SaveEditsAsync();
                });

                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        // 合并图斑
        private void MergeFeature(string targetName, string bjField, long bjID, double miniAngle, string field_group)
        {
            // 初始化要合并的要素OID
            List<long> objectIDList = new List<long>();
            // OID字段
            string oidField = targetName.TargetIDFieldName();


            // 获取图层的要素类
            FeatureClass featureClass = targetName.TargetFeatureClass();
            FeatureLayer featureLayer = targetName.TargetFeatureLayer();

            // 遍历面要素类中的所有要素
            var queryFilter = new QueryFilter();
            queryFilter.WhereClause = $"{bjField} = '{bjID}'";

            RowCursor rowCursor = featureClass.Search(queryFilter);

            // 遍历图层的要素
            while (rowCursor.MoveNext())
            {
                using Feature feature = (Feature)rowCursor.Current;
                
                // 分组字段值
                string group1 = field_group == "" ? "" : feature[field_group]?.ToString();
                
                // 标记字段
                string bj = feature[bjField]?.ToString();

                // 获取要素的几何
                Polygon polygon = feature.GetShape() as Polygon;

                // 检查角度
                // 获取几何的所有段
                var parts = polygon.Parts.ToList();

                foreach (ReadOnlySegmentCollection collection in parts)
                {
                    // 初始化点集
                    List<MapPoint> points = new List<MapPoint>();
                    // 每个环进行处理（第一个为外环，其它为内环）
                    foreach (Segment segment in collection)
                    {
                        MapPoint mapPoint = segment.StartPoint;     // 获取点
                        points.Add(mapPoint);
                    }
                    // 补充首末点，方便下一步计算
                    points.Add(points[0]);
                    points.Insert(0, points[points.Count - 2]);

                    // 计算角度
                    for (int i = 1; i < points.Count - 1; i++)
                    {
                        // 当前点及前后两个点
                        MapPoint p1 = points[i - 1];
                        MapPoint p2 = points[i];
                        MapPoint p3 = points[i + 1];

                        // 获取三点形成的夹角
                        double angle = GetAngle(p1, p2, p3);

                        // 符合条件的
                        if (angle < miniAngle)
                        {
                            // 存储要合并的图斑ID
                            objectIDList.Add(long.Parse(feature[oidField].ToString()));

                            // 创建空间查询过滤器，以获取与源要素有重叠的目标要素
                            SpatialQueryFilter spatialFilter = new SpatialQueryFilter
                            {
                                FilterGeometry = polygon,
                                SpatialRelationship = SpatialRelationship.Intersects
                            };

                            // 在目标图层中查询与源要素重叠的要素
                            using RowCursor identityCursor = featureClass.Search(spatialFilter);

                            // 标记一个面积，寻找最大面积的图斑
                            double maxArea = 0;
                            long initOID = 0;

                            while (identityCursor.MoveNext())
                            {
                                using Feature identityFeature = (Feature)identityCursor.Current;

                                // 分组字段值
                                string group2 = field_group == "" ? "" : identityFeature[field_group]?.ToString();

                                // 标记字段
                                string bj_iden = identityFeature[bjField]?.ToString();
                                // OID字段
                                long oid_iden = long.Parse(identityFeature[oidField]?.ToString());

                                // 获取目标要素的几何
                                Polygon identityGeometry = identityFeature.GetShape() as Polygon;

                                // 面积
                                double area = identityGeometry.Area;

                                // 计算源要素与目标要素的重叠面积
                                Polygon intersection = GeometryEngine.Instance.Intersection(identityGeometry, polygon) as Polygon;

                                // 如果存在相交，则进行下一步处理
                                if (intersection == null) { continue; }   //  无相交，则跳过

                                if (field_group != "" && group1 != group2) { continue; }   // 字段值不同，也跳过

                                if (bj_iden != bj && area > maxArea)   // 标记字段不相等的情况下，且面积更大的时候，更新OID
                                {
                                    maxArea = area;    // 更新maxArea
                                    initOID = oid_iden;   // 更新initOID
                                }
                            }
                            objectIDList.Add(initOID);
                        }
                    }
                }

            }

            // 集合倒序一下，不然合并顺序有误
            objectIDList.Reverse();

            // 合并
            if (objectIDList.Count == 2)
            {
                GisTool.MergeFeatures(featureLayer, objectIDList);
            }
        }


        // 三点判断角度
        private static double GetAngle(MapPoint p1, MapPoint p2, MapPoint p3)
        {
            var angle1 = Math.Atan2(p2.Y - p1.Y, p2.X - p1.X);
            var angle2 = Math.Atan2(p3.Y - p2.Y, p3.X - p2.X);

            var angle = Math.Abs((angle2 - angle1) * 180 / Math.PI);

            // 确保角度在0到180度之间
            if (angle > 180)
                angle = 360 - angle;

            return Math.Abs(180 - angle);
        }

        // 计算从点p1到p2的距离d后的位置
        private static MapPoint CalPosition(MapPoint p1, MapPoint p2, double distance, SpatialReference sr)
        {
            // p1到p2的距离
            double d_12 = Math.Sqrt(Math.Pow((p2.X - p1.X), 2) + Math.Pow((p2.Y - p1.Y), 2));

            // 计算x方向上的移动距离
            double d_x = (distance * (p2.X - p1.X)) / d_12;
            // 计算y方向上的移动距离
            double d_y = (distance * (p2.Y - p1.Y)) / d_12;

            // 移动后的x坐标
            double x = p1.X + d_x;
            // 移动后的y坐标
            double y = p1.Y + d_y;

            MapPoint mapPoint = MapPointBuilder.CreateMapPoint(x, y, sr);

            return mapPoint;
        }


        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/145978114";
            UITool.Link2Web(url);
        }

        private void combox_field_group_DropOpen(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field_group);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }
    }
}
