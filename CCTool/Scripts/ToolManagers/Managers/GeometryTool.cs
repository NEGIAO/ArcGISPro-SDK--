using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.UI.ProMapTool;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Managers
{
    public class GeometryTool
    {
        // 判断两点的方位关系【东，西，南，北】
        public static string Get4Direction(MapPoint p1, MapPoint p2)
        {
            string result = "";

            double deltaX = p2.X - p1.X;
            double deltaY = p2.Y - p1.Y;

            // 判断是否为同一点
            if (Math.Abs(deltaX) < 0.001 && Math.Abs(deltaY) < 0.001)
            {
                result =  "重叠";
            }

            // 比较X/Y轴偏移量的绝对值大小
            if (Math.Abs(deltaX) > Math.Abs(deltaY))
            {
                // X轴偏移量更大时，判断东/西
                result = deltaX > 0 ? "东面" : "西面";
            }
            else
            {
                // Y轴偏移量更大时，判断北/南
                result = deltaY > 0 ? "北面" : "南面";
            }

            return result;
        }

        // 判断两点的方位关系【东，西，南，北，北偏东、北偏西、南偏东、南偏西】
        public static string Get8Direction(MapPoint p1, MapPoint p2)
        {
            string result = "";

            double deltaX = p2.X - p1.X;
            double deltaY = p2.Y - p1.Y;

            // 判断是否为同一点
            if (Math.Abs(deltaX) < 0.001 && Math.Abs(deltaY) < 0.001)
            {
                result = "重叠";
            }

            // 计算方位角（北方为0度，顺时针0-360度）
            double angle = Math.Atan2(deltaX, deltaY) * 180 / Math.PI;
            if (angle < 0) angle += 360;

            // 判断8个方位（正方向覆盖2度）
            if (angle >= 359 || angle < 1) result = "正北";
            else if (angle >= 1 && angle < 89) result = "北偏东";
            else if (angle >= 89 && angle < 91) result = "正东";
            else if (angle >= 91 && angle < 179) result = "南偏东";
            else if (angle >= 179 && angle < 181) result = "正南";
            else if (angle >= 181 && angle < 269) result = "南偏西";
            else if (angle >= 269 && angle < 271) result = "正西";
            else result = "北偏西"; // 271-359度

            return result;
        }

        // 判断两点的方位关系【东，西，南，北，北偏东、北偏西、南偏东、南偏西】
        public static string Get8Direction(double xA, double yA, double xB, double yB)
        {
            string result = "";

            double deltaX = xB - xA;
            double deltaY = yB - yA;

            // 判断是否为同一点
            if (Math.Abs(deltaX) < 0.001 && Math.Abs(deltaY) < 0.001)
            {
                result = "重叠";
            }

            // 计算方位角（北方为0度，顺时针0-360度）
            double angle = Math.Atan2(deltaX, deltaY) * 180 / Math.PI;
            if (angle < 0) angle += 360;

            // 判断8个方位（正方向覆盖2度）
            if (angle >= 337.5 || angle < 22.5)  result = "北";
            else if (angle >= 22.5 && angle < 67.5) result = "北偏东";
            else if (angle >= 67.5 && angle < 112.5) result = "东";
            else if (angle >= 112.5 && angle < 157.5) result = "南偏东";
            else if (angle >= 157.5 && angle < 202.5) result = "南";
            else if (angle >= 202.5 && angle < 247.5) result = "南偏西";
            else if (angle >= 247.5 && angle < 292.5) result = "西";
            else result = "北偏西";

            return result;
        }

        // 相邻线连接
        public static void MergeNearLine(FeatureLayer lineLayer, double minDistance)
        {
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

            // 3. 查找最近端点并创建新线段
            var newLines = new List<Polyline>();
            var processedPairs = new HashSet<(long, long)>();

            foreach (var ep in endpoints)
            {
                // 获取最近且满足距离要求的候选点
                var candidates = endpoints
                    .Where(e => e.ParentOID != ep.ParentOID)
                    .Select(e => new {
                        Point = e.Point,
                        ParentOID = e.ParentOID,
                        Distance = GeometryEngine.Instance.Distance(ep.Point, e.Point) // 计算距离
                    })
                    .Where(e => e.Distance <= minDistance) // 新增距离过滤
                    .OrderBy(e => e.Distance)
                    .FirstOrDefault();

                if (candidates == null) continue;

                if (candidates.Point != null && !processedPairs.Contains((ep.ParentOID, candidates.ParentOID)) && !processedPairs.Contains((candidates.ParentOID, ep.ParentOID)))
                {
                    // 创建新线段
                    var newLine = PolylineBuilderEx.CreatePolyline(
                        new[] { ep.Point, candidates.Point },
                        lineLayer.GetSpatialReference());

                    newLines.Add(newLine);
                    processedPairs.Add((ep.ParentOID, candidates.ParentOID));

                    // 4. 连接原始线段和新线段
                    var originalLine1 = polylines.First(f => f.OID == ep.ParentOID).polyline;
                    var originalLine2 = polylines.First(f => f.OID == candidates.ParentOID).polyline;

                    // 合并三条线段
                    var merged = PolylineBuilderEx.CreatePolyline(
                        new[]
                        {
                                    originalLine1,
                                    newLine,
                                    originalLine2
                        });

                    // 将合并后的几何更新到要素（示例仅更新第一条要素）
                    var editOperation = new EditOperation();
                    editOperation.Modify(lineLayer, ep.ParentOID, merged);
                    editOperation.Delete(lineLayer, candidates.ParentOID);
                    editOperation.Execute();
                }
            }

            // 保存编辑 
            Project.Current.SaveEditsAsync();
        }

        // 闭合线
        public static void CloseLine(string gdbPath, string fcName, double minDistance)
        {
            // 获取复制后的要素
            using Geodatabase geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));
            FeatureClass featureClass = geodatabase.OpenDataset<FeatureClass>(fcName);

            // 获取要素图层的编辑操作
            var editOperation = new EditOperation();
            using (RowCursor rowCursor = featureClass.Search())
            {
                while (rowCursor.MoveNext())
                {
                    using Feature feature = rowCursor.Current as Feature;      // 获取要素
                    // 获取要素的几何对象
                    Polyline polyline = feature.GetShape() as Polyline;
                    // 获取线要素的起点和终点
                    MapPoint start_point = polyline.Points.First();
                    MapPoint end_point = polyline.Points.Last();
                    // 获取起始点的距离
                    double distnace = GeometryEngine.Instance.Distance(start_point, end_point);
                    // 如果起始点的距离超过限界，就跳过
                    if (distnace>minDistance)
                    {
                        continue;
                    }

                    // 如果起始点和终止点的坐标不同，即线不闭合，则进行线闭合操作
                    if (!start_point.Coordinate2D.Equals(end_point.Coordinate2D))
                    {
                        List<Coordinate2D> pts = new List<Coordinate2D>();
                        // 将第一个部分闭合
                        foreach (var part in polyline.Points)
                        {
                            pts.Add(part.Coordinate2D);
                        }
                        pts.Add(start_point.Coordinate2D);
                        // 创建 PolylineBuilder 对象并闭合线要素
                        var builder = new PolylineBuilder(pts);
                        // 获取闭合后的几何
                        var closedGeometry = builder.ToGeometry();

                        // 设置要素的几何
                        feature.SetShape(closedGeometry);
                    }
                    feature.Store();
                }
            }
            // 执行编辑
            editOperation.Execute();

            // 保存编辑 
            Project.Current.SaveEditsAsync();
        }

    }
}
