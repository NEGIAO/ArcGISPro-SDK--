using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Data.Raster;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Extensions
{
    public static class GisExtension
    {
        // 重排界址点，内外环都按逆时针
        public static List<List<MapPoint>> ReshotMapPointWise(this Polygon polygon)
        {
            List<List<MapPoint>> result = new List<List<MapPoint>>();

            // 获取面要素的所有点（内外环）
            List<List<MapPoint>> mapPoints = polygon.MapPointsFromPolygon();

            // 每个环进行处理
            for (int j = 0; j < mapPoints.Count; j++)
            {
                List<MapPoint> vertices = mapPoints[j];

                // 判断顺逆时针，如果有问题就调整反向
                bool isClockwise = vertices.IsColckwise();
                if (isClockwise)    // 如果是顺时针，则调整反向
                {
                    vertices.ReversePoint();
                }

                // 在末尾加起始点
                vertices.Add(vertices[0]);

                result.Add(vertices);
            }
            // 返回值
            return result;
        }

        // 线要素重排界址点，内外环顺逆时针可选
        public static List<List<MapPoint>> ReshotPolylineMapPointWise(this Polyline polyline, bool isWise = false)
        {
            List<List<MapPoint>> result = new List<List<MapPoint>>();

            // 获取面要素的所有点（内外环）
            List<List<MapPoint>> mapPoints = polyline.MapPointsFromPolyline();

            // 每个环进行处理
            for (int j = 0; j < mapPoints.Count; j++)
            {
                List<MapPoint> vertices = mapPoints[j];

                // 判断顺逆时针，如果有问题就调整反向
                bool isClockwise = vertices.IsColckwise();
                if (isClockwise != isWise)
                {
                    vertices.ReversePoint();
                }

                // 在末尾加起始点
                vertices.Add(vertices[0]);

                result.Add(vertices);
            }
            // 返回值
            return result;
        }

        // 重排界址点，从西北角开始，并理清顺逆时针
        public static List<List<MapPoint>> ReshotMapPoint(this Polygon polygon, bool isAddFirst = true)
        {
            List<List<MapPoint>> result = new List<List<MapPoint>>();

            // 获取面要素的所有点（内外环）
            List<List<MapPoint>> mapPoints = polygon.MapPointsFromPolygon();
            // 获取每个环的最西北点
            List<List<double>> NWPoints = polygon.NWPointsFromPolygon();

            // 每个环进行处理
            for (int j = 0; j < mapPoints.Count; j++)
            {
                List<MapPoint> newVertices = new List<MapPoint>();
                List<MapPoint> vertices = mapPoints[j];
                // 获取要素的最西北点坐标
                double XMin = NWPoints[j][0];
                double YMax = NWPoints[j][1];

                // 找出西北点【离西北角（Xmin,Ymax）最近的点】
                int targetIndex = 0;
                double maxDistance = 10000000;
                for (int i = 0; i < vertices.Count; i++)
                {
                    // 计算和西北角的距离
                    double distance = Math.Sqrt(Math.Pow(vertices[i].X - XMin, 2) + Math.Pow(vertices[i].Y - YMax, 2));
                    // 如果小于上一个值，则保存新值，直到找出最近的点
                    if (distance < maxDistance)
                    {
                        targetIndex = i;
                        maxDistance = distance;
                    }
                }

                // 根据最近点重排顺序
                newVertices = vertices.GetRange(targetIndex, vertices.Count - targetIndex);
                vertices.RemoveRange(targetIndex, vertices.Count - targetIndex);
                newVertices.AddRange(vertices);

                // 判断顺逆时针，如果有问题就调整反向
                bool isClockwise = newVertices.IsColckwise();
                if (!isClockwise && j == 0)    // 如果是外环，且逆时针，则调整反向
                {
                    newVertices.ReversePoint();
                }
                if (isClockwise && j > 0)    // 如果是内环，且顺时针，则调整反向
                {
                    newVertices.ReversePoint();
                }

                // 在末尾加起始点
                if (isAddFirst)
                {
                    newVertices.Add(newVertices[0]);
                }

                result.Add(newVertices);
            }
            // 返回值
            return result;
        }

        // 重排界址点，自定义起始点，并理清顺逆时针
        public static List<List<MapPoint>> ReshotMapPointByCustom(this Polygon polygon, List<double> xy)
        {
            List<List<MapPoint>> result = new List<List<MapPoint>>();

            // 获取面要素的所有点（内外环）
            List<List<MapPoint>> mapPoints = polygon.MapPointsFromPolygon();

            // 每个环进行处理
            for (int j = 0; j < mapPoints.Count; j++)
            {
                List<MapPoint> newVertices = new List<MapPoint>();
                List<MapPoint> vertices = mapPoints[j];
                // 获取要素的自定义点坐标
                double XMin = xy[0];
                double YMax = xy[1];

                // 找出xy最近的点
                int targetIndex = 0;
                double maxDistance = 10000000;
                for (int i = 0; i < vertices.Count; i++)
                {
                    // 计算和西北角的距离
                    double distance = Math.Sqrt(Math.Pow(vertices[i].X - XMin, 2) + Math.Pow(vertices[i].Y - YMax, 2));
                    // 如果小于上一个值，则保存新值，直到找出最近的点
                    if (distance < maxDistance)
                    {
                        targetIndex = i;
                        maxDistance = distance;
                    }
                }

                // 根据最近点重排顺序
                newVertices = vertices.GetRange(targetIndex, vertices.Count - targetIndex);
                vertices.RemoveRange(targetIndex, vertices.Count - targetIndex);
                newVertices.AddRange(vertices);

                // 判断顺逆时针，如果有问题就调整反向
                bool isClockwise = newVertices.IsColckwise();
                if (!isClockwise && j == 0)    // 如果是外环，且逆时针，则调整反向
                {
                    newVertices.ReversePoint();
                }
                if (isClockwise && j > 0)    // 如果是内环，且顺时针，则调整反向
                {
                    newVertices.ReversePoint();
                }

                // 在末尾加起始点
                newVertices.Add(newVertices[0]);

                result.Add(newVertices);
            }
            // 返回值
            return result;
        }

        // 重排界址点，自定义起始点，并理清顺逆时针
        public static List<List<MapPoint>> ReshotPolylineMapPointByCustom(this Polyline polyline, List<double> xy)
        {
            List<List<MapPoint>> result = new List<List<MapPoint>>();

            // 获取面要素的所有点（内外环）
            List<List<MapPoint>> mapPoints = polyline.MapPointsFromPolyline();

            // 每个环进行处理
            for (int j = 0; j < mapPoints.Count; j++)
            {
                List<MapPoint> newVertices = new List<MapPoint>();
                List<MapPoint> vertices = mapPoints[j];
                // 获取要素的自定义点坐标
                double XMin = xy[0];
                double YMax = xy[1];

                // 找出xy最近的点
                int targetIndex = 0;
                double maxDistance = 10000000;
                for (int i = 0; i < vertices.Count; i++)
                {
                    // 计算和西北角的距离
                    double distance = Math.Sqrt(Math.Pow(vertices[i].X - XMin, 2) + Math.Pow(vertices[i].Y - YMax, 2));
                    // 如果小于上一个值，则保存新值，直到找出最近的点
                    if (distance < maxDistance)
                    {
                        targetIndex = i;
                        maxDistance = distance;
                    }
                }

                // 根据最近点重排顺序
                newVertices = vertices.GetRange(targetIndex, vertices.Count - targetIndex);
                vertices.RemoveRange(targetIndex, vertices.Count - targetIndex);
                newVertices.AddRange(vertices);

                // 判断顺逆时针，如果有问题就调整反向
                bool isClockwise = newVertices.IsColckwise();
                if (!isClockwise && j == 0)    // 如果是外环，且逆时针，则调整反向
                {
                    newVertices.ReversePoint();
                }
                if (isClockwise && j > 0)    // 如果是内环，且顺时针，则调整反向
                {
                    newVertices.ReversePoint();
                }

                // 在末尾加起始点
                newVertices.Add(newVertices[0]);

                result.Add(newVertices);
            }
            // 返回值
            return result;
        }

        // 重排界址点，从西北角开始，并理清顺逆时针，并返回面要素
        public static Polygon ReshotMapPointReturnPolygon(this Polygon polygon)
        {
            List<List<MapPoint>> mapPoints = polygon.ReshotMapPoint();

            // 判断是否有内环，如果没有内环
            if (mapPoints.Count == 1)
            {
                Polygon resultPolygon = PolygonBuilderEx.CreatePolygon(mapPoints[0]);
                // 返回值
                return resultPolygon;
            }
            // 如果有内环
            else
            {
                // 生成外环点集
                List<Coordinate2D> outerPts = new List<Coordinate2D>();
                foreach (MapPoint pt in mapPoints[0])
                {
                    outerPts.Add(new Coordinate2D(pt.X, pt.Y));
                }
                // 创建外环面
                PolygonBuilderEx pb = new PolygonBuilderEx(outerPts);

                // 移除外环，剩下内环进行处理
                mapPoints.RemoveAt(0);
                // 收集内环集合
                foreach (List<MapPoint> innerPoints in mapPoints)
                {
                    // 获取内环点集
                    List<Coordinate2D> innerPts = new List<Coordinate2D>();
                    foreach (MapPoint pt in innerPoints)
                    {
                        innerPts.Add(new Coordinate2D(pt.X, pt.Y));
                    }
                    // 将内环点集加入到外环面处理
                    pb.AddPart(innerPts);
                }
                // 返回最终的面
                return pb.ToGeometry();
            }
        }

        // 设置线要素的顺逆时针
        public static void SetPolylineWise(this string in_line, bool isWise)
        {
            // 获取目标FeatureLayer
            FeatureLayer featurelayer = in_line.TargetFeatureLayer();

            // 遍历面要素类中的所有要素
            RowCursor cursor = featurelayer.Search();
            while (cursor.MoveNext())
            {
                using var feature = cursor.Current as Feature;
                // 获取要素的几何
                Polyline geometry = feature.GetShape() as Polyline;
                if (geometry != null)
                {
                    // 面要素的所有折点进行重排【顺逆时针重排】
                    List<List<MapPoint>> mpts = geometry.ReshotPolylineMapPointWise(isWise);
                    Polyline resultPolyline = PolylineBuilderEx.CreatePolyline(mpts[0]);
                    // 重新设置要素并保存
                    feature.SetShape(resultPolyline);
                    feature.Store();
                }
            }
        }

        // 重排界址点，自定义起始点，并理清顺逆时针，并返回线要素
        public static Polyline ReshotMapPointReturnPolylineByCustom(this Polyline polyline, List<double> xy)
        {
            List<List<MapPoint>> mapPoints = polyline.ReshotPolylineMapPointByCustom(xy);

            Polyline resultPolyline = PolylineBuilderEx.CreatePolyline(mapPoints[0]);
            // 返回值
            return resultPolyline;
        }

        // 重排界址点，自定义起始点，并理清顺逆时针，并返回面要素
        public static Polygon ReshotMapPointReturnPolygonByCustom(this Polygon polygon, List<double> xy)
        {
            List<List<MapPoint>> mapPoints = polygon.ReshotMapPointByCustom(xy);

            // 判断是否有内环，如果没有内环
            if (mapPoints.Count == 1)
            {
                Polygon resultPolygon = PolygonBuilderEx.CreatePolygon(mapPoints[0]);
                // 返回值
                return resultPolygon;
            }
            // 如果有内环
            else
            {
                // 生成外环点集
                List<Coordinate2D> outerPts = new List<Coordinate2D>();
                foreach (MapPoint pt in mapPoints[0])
                {
                    outerPts.Add(new Coordinate2D(pt.X, pt.Y));
                }
                // 创建外环面
                PolygonBuilderEx pb = new PolygonBuilderEx(outerPts);

                // 移除外环，剩下内环进行处理
                mapPoints.RemoveAt(0);
                // 收集内环集合
                foreach (List<MapPoint> innerPoints in mapPoints)
                {
                    // 获取内环点集
                    List<Coordinate2D> innerPts = new List<Coordinate2D>();
                    foreach (MapPoint pt in innerPoints)
                    {
                        innerPts.Add(new Coordinate2D(pt.X, pt.Y));
                    }
                    // 将内环点集加入到外环面处理
                    pb.AddPart(innerPts);
                }
                // 返回最终的面
                return pb.ToGeometry();
            }
        }

        // 获取面要素的所有点【isAddFirst为true时，把第一个点加到末尾】
        public static List<List<MapPoint>> MapPointsFromPolygon(this Polygon polygon, bool isAddFirst = false)
        {
            List<List<MapPoint>> mapPoints = new List<List<MapPoint>>();

            // 获取面要素的部件（内外环）
            var parts = polygon.Parts.ToList();
            foreach (ReadOnlySegmentCollection collection in parts)
            {
                List<MapPoint> points = new List<MapPoint>();
                // 每个环进行处理（第一个为外环，其它为内环）
                foreach (Segment segment in collection)
                {
                    MapPoint mapPoint = segment.StartPoint;     // 获取点
                    points.Add(mapPoint);
                }
                // 是否追加第一个点
                if (isAddFirst)
                {
                    points.Add(points[0]);
                }

                mapPoints.Add(points);
            }
            return mapPoints;
        }


        // 获取线要素的所有点【isAddFirst为true时，把第一个点加到末尾】
        public static List<List<MapPoint>> MapPointsFromPolyline(this Polyline polyline, bool isAddFirst = false)
        {
            List<List<MapPoint>> mapPoints = new List<List<MapPoint>>();

            // 获取线要素的部件
            var parts = polyline.Parts.ToList();
            foreach (ReadOnlySegmentCollection collection in parts)
            {
                List<MapPoint> points = new List<MapPoint>();
                // 每个环进行处理（第一个为外环，其它为内环）
                foreach (Segment segment in collection)
                {
                    MapPoint mapPoint = segment.StartPoint;     // 获取点
                    points.Add(mapPoint);
                }
                // 是否追加第一个点
                if (isAddFirst)
                {
                    points.Add(points[0]);
                }

                mapPoints.Add(points);
            }
            return mapPoints;
        }


        // 获取面要素的所有环的最西北点坐标
        public static List<List<double>> NWPointsFromPolygon(this Polygon polygon)
        {
            List<List<double>> NWPoints = new List<List<double>>();

            // 获取面要素的部件（内外环）
            var parts = polygon.Parts;
            foreach (ReadOnlySegmentCollection collection in parts)
            {
                List<double> point = new List<double>();
                // 初始化西北角的点
                double XMin = 100000000;
                double YMax = 0;

                // 每个环进行处理（第一个为外环，其它为内环）
                foreach (Segment segment in collection)
                {
                    MapPoint mapPoint = segment.StartPoint;     // 获取点
                    if (mapPoint.X < XMin)
                    {
                        XMin = mapPoint.X;
                    }
                    if (mapPoint.Y > YMax)
                    {
                        YMax = mapPoint.Y;
                    }
                }
                point.Add(XMin);
                point.Add(YMax);
                NWPoints.Add(point);
            }
            return NWPoints;
        }

        // 判断点集合形成的面的面积为正值还是负值，用以判断是顺时针还是逆时针
        public static bool IsColckwise(this List<MapPoint> newVertices)
        {
            // 判断界址点是否顺时针排序
            double x1, y1, x2, y2;
            double area = 0;
            for (int i = 0; i < newVertices.Count; i++)
            {
                x1 = newVertices[i].X;
                y1 = newVertices[i].Y;
                if (i == newVertices.Count - 1)
                {
                    x2 = newVertices[0].X;
                    y2 = newVertices[0].Y;
                }
                else
                {
                    x2 = newVertices[i + 1].X;
                    y2 = newVertices[i + 1].Y;
                }
                area += x1 * y2 - x2 * y1;
            }
            if (area > 0)     // 逆时针
            {
                return false;
            }
            else        // 顺时针
            {
                return true;
            }
        }

        // MapPoint第一个点不变，反方向排列
        public static void ReversePoint(this List<MapPoint> newVertices)
        {
            newVertices.Reverse();
            newVertices.Insert(0, newVertices[newVertices.Count - 1]);
            newVertices.RemoveAt(newVertices.Count - 1);
        }

        // 判断是否有Z值
        public static bool IsHasZ(this string target_fc)
        {
            bool hasZ = false;

            FeatureClass featureClass = target_fc.TargetFeatureClass();

            // 遍历面要素类中的所有要素
            RowCursor cursor = featureClass.Search();
            while (cursor.MoveNext())
            {
                using var feature = cursor.Current as Feature;
                // 获取要素的几何
                Polygon geometry = feature.GetShape() as Polygon;

                if (geometry.HasZ)
                {
                    hasZ = true;
                    break;
                }
                else
                {
                    hasZ = false;
                    break;
                }
            }
            return hasZ;
        }

        // 判断要素数据集是否存在
        public static bool IsHaveDataset(this string gdb_path, string dt_name)
        {
            bool isHaveGDB = Directory.Exists(gdb_path);
            if (!isHaveGDB)
            {
                return false;
            }
            else
            {
                // 打开地理数据库
                using var geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_path)));
                // 获取所有要素数据集的名称
                var datasets = geodatabase.GetDefinitions<FeatureDatasetDefinition>();
                // 检查要素数据集是否存在
                var exists = datasets.Any(datasetName => datasetName.GetName().Equals(dt_name));
                return exists;
            }
        }

        // 判断要素类是否存在
        public static bool IsHaveFeaturClass(this string gdb_path, string fc_name)
        {
            bool isHaveGDB = Directory.Exists(gdb_path);
            if (!isHaveGDB)
            {
                return false;
            }
            else
            {
                // 打开地理数据库
                using var geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_path)));
                // 获取所有要素数据集的名称
                var fcs = geodatabase.GetDefinitions<FeatureClassDefinition>();
                // 检查要素数据集是否存在
                var exists = fcs.Any(datasetName => datasetName.GetName().Equals(fc_name));
                return exists;
            }
        }

        // 判断独立表是否存在
        public static bool IsHaveStandaloneTable(this string gdb_path, string tb_name)
        {
            bool isHaveGDB = Directory.Exists(gdb_path);
            if (!isHaveGDB)
            {
                return false;
            }
            else
            {
                // 打开地理数据库
                using var geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_path)));
                // 获取所有要素数据集的名称
                var tables = geodatabase.GetDefinitions<TableDefinition>();
                // 检查要素数据集是否存在
                var exists = tables.Any(datasetName => datasetName.GetName().Equals(tb_name));
                return exists;
            }
        }

        // 获取GDB数据库里的所有FeatureClass路径
        public static List<string> GetFeatureClassPathFromGDB(this string gdbPath)
        {
            List<string> result = new List<string>();
            // 打开GDB数据库
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));
            // 获取所有要素类
            IReadOnlyList<FeatureClassDefinition> featureClasses = gdb.GetDefinitions<FeatureClassDefinition>();
            foreach (FeatureClassDefinition featureClass in featureClasses)
            {
                using (FeatureClass fc = gdb.OpenDataset<FeatureClass>(featureClass.GetName()))
                {
                    // 获取要素类路径
                    string fc_path = fc.GetPath().ToString().Replace("file:///", "").Replace("/", @"\");
                    result.Add(fc_path);
                }
            }
            return result;
        }

        // 获取GDB数据库里的所有表格路径
        public static List<string> GetStandaloneTablePathFromGDB(this string gdbPath)
        {
            List<string> result = new List<string>();
            // 打开GDB数据库
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));
            // 获取所有独立表
            IReadOnlyList<TableDefinition> tables = gdb.GetDefinitions<TableDefinition>();
            foreach (TableDefinition tableDef in tables)
            {
                using (Table table = gdb.OpenDataset<Table>(tableDef.GetName()))
                {
                    // 获取要素类路径
                    string fc_path = table.GetPath().ToString().Replace("file:///", "").Replace("/", @"\");
                    result.Add(fc_path);
                }
            }
            return result;
        }

        // 获取数据库下的所有要素类和独立表的完整路径
        public static List<string> GetFeatureClassAndTablePath(this string gdbPath)
        {
            // 获取要素类的完整路径
            List<string> fcs = gdbPath.GetFeatureClassPathFromGDB();
            // 获取独立表的完整路径
            List<string> tbs = gdbPath.GetStandaloneTablePathFromGDB();
            // 合并列表
            fcs.AddRange(tbs);
            return fcs;
        }

        // 获取数据库下的所有要素类的名称
        public static List<string> GetFeatureClassNameFromGDB(this string gdb_path)
        {
            List<string> result = new List<string>();
            // 打开GDB数据库
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_path)));
            // 获取所有要素类
            IReadOnlyList<FeatureClassDefinition> featureClasses = gdb.GetDefinitions<FeatureClassDefinition>();
            foreach (FeatureClassDefinition featureClass in featureClasses)
            {
                string fc_name = featureClass.GetName();
                result.Add(fc_name);
            }

            return result;
        }

        // 获取数据库下的所有独立表的名称
        public static List<string> GetTableName(this string gdb_path)
        {
            List<string> result = new List<string>();
            // 打开GDB数据库
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_path)));
            // 获取所有独立表
            IReadOnlyList<TableDefinition> tables = gdb.GetDefinitions<TableDefinition>();
            foreach (TableDefinition tableDef in tables)
            {
                string tb_name = tableDef.GetName();
                result.Add(tb_name);
            }
            return result;
        }

        // 获取数据库下的所有要素类和独立表的名称
        public static List<string> GetFeatureClassAndTableName(this string gdb_path)
        {
            // 获取要素类的完整路径
            List<string> fcs = gdb_path.GetFeatureClassNameFromGDB();
            // 获取独立表的完整路径
            List<string> tbs = gdb_path.GetTableName();
            // 合并列表
            fcs.AddRange(tbs);
            return fcs;
        }

        // 获取数据库下的所有栅格的完整路径
        public static List<string> GetRasterPath(this string gdb_path)
        {
            List<string> result = new List<string>();
            // 打开GDB数据库
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_path)));
            // 获取所有要素类
            IReadOnlyList<RasterDatasetDefinition> rasterDefinitions = gdb.GetDefinitions<RasterDatasetDefinition>();
            foreach (RasterDatasetDefinition rasterDefinition in rasterDefinitions)
            {
                using (RasterDataset rd = gdb.OpenDataset<RasterDataset>(rasterDefinition.GetName()))
                {
                    // 获取要素类路径
                    string rd_path = rd.GetPath().ToString().Replace("file:///", "").Replace("/", @"\");
                    result.Add(rd_path);
                }
            }
            return result;
        }

        // 获取数据库下的所有要素数据集的完整路径
        public static List<string> GetDataBasePath(this string gdb_path)
        {
            List<string> result = new List<string>();
            // 打开GDB数据库
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_path)));
            // 获取所有要素类
            IReadOnlyList<FeatureDatasetDefinition> featureDatases = gdb.GetDefinitions<FeatureDatasetDefinition>();
            foreach (FeatureDatasetDefinition featureDatase in featureDatases)
            {
                using (FeatureDataset fd = gdb.OpenDataset<FeatureDataset>(featureDatase.GetName()))
                {
                    // 获取要素类路径
                    string fd_path = fd.GetPath().ToString().Replace("file:///", "").Replace("/", @"\");
                    result.Add(fd_path);
                }
            }

            return result;
        }

        // 获取数据库下的所有要素数据集的名称
        public static List<string> GetDataBaseName(this string gdb_path)
        {
            List<string> result = new List<string>();
            // 打开GDB数据库
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_path)));
            // 获取所有要素类
            IReadOnlyList<FeatureDatasetDefinition> featureDatases = gdb.GetDefinitions<FeatureDatasetDefinition>();
            foreach (FeatureDatasetDefinition featureDatase in featureDatases)
            {
                // 获取要素类路径
                string fd_name = featureDatase.GetName();
                result.Add(fd_name);
            }

            return result;
        }

        // 获取独立图层名
        public static string GetLayerSingleName(this string fullPath)
        {
            string singleName = fullPath;
            // 如果是多层次的，获取最后一层
            if (fullPath.Contains(@"\"))
            {
                singleName = fullPath[(fullPath.LastIndexOf(@"\") + 1)..];
            }
            return singleName;
        }

        // 获取地图中的所有要素图层【带图层结构】【字典】
        public static Dictionary<FeatureLayer, string> AllFeatureLayersDic(this Map map)
        {
            Dictionary<FeatureLayer, string> dic = new Dictionary<FeatureLayer, string>();
            List<string> layers = new List<string>();
            List<FeatureLayer> lys = new List<FeatureLayer>();
            // 获取所有要素图层
            List<FeatureLayer> featureLayers = map.GetLayersAsFlattenedList().OfType<FeatureLayer>().ToList();

            foreach (FeatureLayer featureLayer in featureLayers)
            {
                if (featureLayer != null)
                {
                    string lyPath = featureLayer.TargetLayerPath();

                    if (lyPath != "")
                    {
                        if (!lyPath.ToLower().Contains(".dwg") && !lyPath.ToLower().Contains(".dxf"))
                        {
                            List<object> list = featureLayer.GetLayerFullName(map, featureLayer.Name);

                            layers.Add((string)list[1]);
                            lys.Add(featureLayer);
                        }
                    }
                }
            }
            // 标记重复
            layers.AddNumbersToDuplicates();
            // 加入字典
            for (int i = 0; i < lys.Count; i++)
            {
                dic.Add(lys[i], layers[i]);
            }
            // 返回值
            return dic;
        }

        // 获取地图中的所有独立表【带图层结构】【字典】
        public static Dictionary<StandaloneTable, string> AllStandaloneTablesDic(this Map map)
        {
            Dictionary<StandaloneTable, string> dic = new Dictionary<StandaloneTable, string>();
            List<string> layers = new List<string>();
            List<StandaloneTable> lys = new List<StandaloneTable>();
            // 获取所有要素图层
            List<StandaloneTable> standaloneTables = map.GetStandaloneTablesAsFlattenedList().ToList();
            foreach (StandaloneTable standaloneTable in standaloneTables)
            {
                List<object> list = standaloneTable.GetLayerFullName(map, standaloneTable.Name);

                layers.Add((string)list[1]);
                lys.Add(standaloneTable);
            }
            // 标记重复
            layers.AddNumbersToDuplicates();
            // 加入字典
            for (int i = 0; i < lys.Count; i++)
            {
                dic.Add(lys[i], layers[i]);
            }
            // 返回值
            return dic;
        }

        // 获取地图中的所有要素图层【带图层结构】
        public static List<string> AllFeatureLayers(this Map map)
        {

            List<string> layers = new List<string>();
            // 获取所有要素图层
            List<FeatureLayer> featureLayers = map.GetLayersAsFlattenedList().OfType<FeatureLayer>().ToList();

            foreach (FeatureLayer featureLayer in featureLayers)
            {
                // 判断下是不是cad图层，cad图层不参与
                if (featureLayer != null)
                {
                    string lyPath = featureLayer.TargetLayerPath();

                    if (lyPath != "")
                    {
                        if (!lyPath.ToLower().Contains(".dwg") && !lyPath.ToLower().Contains(".dxf"))
                        {
                            List<object> list = featureLayer.GetLayerFullName(map, featureLayer.Name);
                            layers.Add((string)list[1]);
                        }
                    }
                }
            }
            return layers.AddNumbersToDuplicates();
        }


        // 获取地图中的所有可见图层【带图层结构】
        public static List<string> CanseeFeatureLayers(this Map map)
        {
            List<string> layers = new List<string>();
            // 获取所有可见图层
            List<FeatureLayer> featureLayers = map.GetLayersAsFlattenedList().OfType<FeatureLayer>().ToList();
            List<RasterLayer> rasterLayers = map.GetLayersAsFlattenedList().OfType<RasterLayer>().ToList();

            if (featureLayers is not null)
            {
                foreach (FeatureLayer featureLayer in featureLayers)
                {
                    List<object> list = featureLayer.GetLayerFullName(map, featureLayer.Name);

                    layers.Add((string)list[1]);
                }
            }

            if (rasterLayers is not null)
            {
                foreach (RasterLayer rasterLayer in rasterLayers)
                {
                    List<object> list = rasterLayer.GetLayerFullName(map, rasterLayer.Name);

                    layers.Add((string)list[1]);
                }
            }

            return layers.AddNumbersToDuplicates();
        }

        // 获取图层的图斑个数
        public static long GetFeatureCount(this FeatureLayer featureLayer)
        {
            // 遍历面要素类中的所有要素
            using var cursor = featureLayer.Search();
            int featureCount = 0;
            while (cursor.MoveNext())
            {
                featureCount++;
            }
            return featureCount;
        }

        // 获取要素类的图斑个数
        public static long GetFeatureCount(this string fcPath)
        {
            // 遍历面要素类中的所有要素
            FeatureClass featureClass = fcPath.TargetFeatureClass();

            if (featureClass is null)
            {
                return 0;
            }
            else
            {
                using var cursor = featureClass.Search();
                int featureCount = 0;
                while (cursor.MoveNext())
                {
                    featureCount++;
                }
                return featureCount;
            }
        }

        // 获取地图中的所有独立表【带图层结构】
        public static List<string> AllStandaloneTables(this Map map)
        {
            List<string> layers = new List<string>();
            // 获取所有要素图层
            List<StandaloneTable> standaloneTables = map.GetStandaloneTablesAsFlattenedList().ToList();
            foreach (StandaloneTable standaloneTable in standaloneTables)
            {
                List<object> list = standaloneTable.GetLayerFullName(map, standaloneTable.Name);

                layers.Add((string)list[1]);
            }
            return layers.AddNumbersToDuplicates();
        }

        // 获取图层的完整名称
        public static List<object> GetLayerFullName(this Object layer, Map map, string lyName)
        {
            List<object> result = new List<object>();
            // 如果是图层
            if (layer is Layer)
            {
                // 如果父对象是Map，直接返回图层名
                if (((Layer)layer).Parent is Map)
                {
                    result.Add(layer);
                    result.Add(lyName);

                    return result;
                }
                else
                {
                    // 如果父对象是不是Map，则找到父对象图层，并循环查找上一个层级
                    Layer paLayer = (Layer)((Layer)layer).Parent;

                    List<object> list = paLayer.GetLayerFullName(map, @$"{paLayer}\{lyName}");

                    return list;
                }
            }
            // 如果是独立表
            else if (layer is StandaloneTable)
            {
                // 如果父对象是Map，直接返回图层名
                if (((StandaloneTable)layer).Parent is Map)
                {
                    result.Add(layer);
                    result.Add(lyName);

                    return result;
                }
                else
                {
                    // 如果父对象是不是Map，则找到父对象图层，并循环查找上一个层级
                    Layer paLayer = (Layer)((StandaloneTable)layer).Parent;

                    List<object> list = paLayer.GetLayerFullName(map, @$"{paLayer}\{lyName}");

                    return list;
                }
            }
            else
            {
                return null;
            }
        }

        // 从图层的完整名称获取图层
        public static Object GetLayerFromFullName(this string layerFullName)
        {
            List<Object> result = new List<object>();

            // 获取当前地图
            Map map = MapView.Active.Map;
            Dictionary<FeatureLayer, string> dicFeatureLayer = map.AllFeatureLayersDic();
            Dictionary<StandaloneTable, string> dicStandaloneTable = map.AllStandaloneTablesDic();
            // 查找要素图层
            foreach (var layer in dicFeatureLayer)
            {
                if (layerFullName == layer.Value)
                {
                    result.Add(layer.Key);
                }
            }

            // 查找独立表
            foreach (var layer in dicStandaloneTable)
            {
                if (layerFullName == layer.Value)
                {
                    result.Add(layer.Key);
                }
            }

            // 返回值
            return result[0];
        }

        //  获取字段的所有唯一值
        public static List<string> GetFieldValues(this object targetPath, string fieldName)
        {
            List<string> fieldValues = new List<string>();
            // 获取Table
            Table table;
            if (targetPath is FeatureLayer layer)
            {
                table = layer.GetTable();
            }
            else
            {
                table = ((string)targetPath).TargetTable();
            }
            // 逐行
            using RowCursor rowCursor = table.Search();
            while (rowCursor.MoveNext())
            {
                using ArcGIS.Core.Data.Row row = rowCursor.Current;
                // 获取value
                var va = row[fieldName];
                if (va != null)
                {
                    string result = va.ToString();
                    // 如果不在列表中，就加入
                    if (!fieldValues.Contains(result))
                    {
                        fieldValues.Add(va.ToString());
                    }
                }
            }
            return fieldValues;
        }

        //  获取字段的所有值
        public static List<string> GetAllFieldValues(this object targetPath, string fieldName)
        {
            List<string> fieldValues = new List<string>();
            // 获取Table
            Table table;
            if (targetPath is FeatureLayer layer)
            {
                table = layer.GetTable();
            }
            else
            {
                table = ((string)targetPath).TargetTable();
            }
            // 逐行找出错误
            using RowCursor rowCursor = table.Search();
            while (rowCursor.MoveNext())
            {
                using ArcGIS.Core.Data.Row row = rowCursor.Current;
                // 获取value
                var va = row[fieldName];
                if (va != null)
                {
                    string result = va.ToString();
                    fieldValues.Add(va.ToString());
                }
            }
            return fieldValues;
        }

        //  获取字段的所有唯一值(含同一值个数)
        public static Dictionary<string, long> GetFieldValuesDic(this object targetPath, string fieldName)
        {
            Dictionary<string, long> dic = new Dictionary<string, long>();
            // 获取Table
            Table table;
            if (targetPath is FeatureLayer layer)
            {
                table = layer.GetTable();
            }
            else
            {
                table = ((string)targetPath).TargetTable();
            }
            // 逐行找出错误
            using RowCursor rowCursor = table.Search();
            while (rowCursor.MoveNext())
            {
                using ArcGIS.Core.Data.Row row = rowCursor.Current;
                // 获取value
                var va = row[fieldName];
                string str = va is null ? "[null]" : va.ToString();
                // 如果不在列表中，就加入
                if (!dic.ContainsKey(str))
                {
                    dic.Add(str, 1);
                }
                // 如果已有，则加数
                else
                {
                    dic[str] += 1;
                }
            }
            return dic;
        }


        //  获取两个字段的值映射表
        public static Dictionary<string, string> Get2FieldValueDic(this object targetPath, string fieldName01, string fieldName02)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            // 获取Table
            Table table;
            if (targetPath is FeatureLayer layer)
            {
                table = layer.GetTable();
            }
            else
            {
                table = ((string)targetPath).TargetTable();
            }
            // 逐行找出
            using RowCursor rowCursor = table.Search();
            while (rowCursor.MoveNext())
            {
                using ArcGIS.Core.Data.Row row = rowCursor.Current;
                // 获取value
                var va01 = row[fieldName01];
                var va02 = row[fieldName02];
                if (va01 != null && va02 != null)
                {
                    string txt01 = va01.ToString();
                    string txt02 = va02.ToString();
                    // 如果不在列表中，就加入
                    if (!dic.ContainsKey(txt01))
                    {
                        dic.Add(txt01, txt02);
                    }
                }
            }
            return dic;
        }

        // 获取要素属性
        public static FeatureClassAtt GetFeatureClassAtt(this string lyName)
        {
            FeatureClassAtt featureClassAtt = new FeatureClassAtt();
            // 获取FeatureClassDescription
            FeatureClassDescription fcd = lyName.TargetFeatureClassDescription();
            // 获取FeatureClass
            FeatureClass featureClass = lyName.TargetFeatureClass();
            // 获取FeatureClass
            FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();

            // 获取属性
            // 名称
            featureClassAtt.Name = fcd.Name;
            // 别名
            featureClassAtt.AliasName = fcd.AliasName;
            // 要素类型
            featureClassAtt.GeometryType = featureClassDefinition.GetShapeType();
            // 是否有Z，M
            featureClassAtt.HasZ = featureClassDefinition.HasZ();
            featureClassAtt.HasM = featureClassDefinition.HasM();
            // 字段定义
            featureClassAtt.FieldDescriptions = fcd.FieldDescriptions;
            // 坐标系
            featureClassAtt.SpatialReference = featureClassDefinition.GetSpatialReference();

            // 要素数量
            featureClassAtt.FeatureCount = featureClass.GetCount();
            // OID字段
            foreach (var item in fcd.FieldDescriptions)
            {
                if (item.FieldType == FieldType.OID)
                {
                    featureClassAtt.OIDField = item.Name;
                    break;
                }
            }

            return featureClassAtt;
        }

        //  获取字段属性
        public static FieldAtt GetFieldAtt(this object targetPath, string fieldName)
        {
            FieldAtt fieldAtt = new FieldAtt();

            List<string> fieldValues = new List<string>();

            Table table;
            if (targetPath is FeatureLayer layer)
            {
                table = layer.GetTable();
            }
            else
            {
                table = ((string)targetPath).TargetTable();
            }

            // 获取字段通用属性
            var fileds = table.GetDefinition().GetFields();
            foreach (Field filed in fileds)
            {
                if (filed.Name == fieldName)
                {
                    fieldAtt.Name = filed.Name;
                    fieldAtt.AliasName = filed.AliasName;
                    fieldAtt.Type = filed.FieldType;
                    fieldAtt.Length = filed.Length;
                    break;
                }
            }

            return fieldAtt;
        }

        // 清除指定Name的Item的
        public static void RemoveItemByName(this StyleProjectItem styleProjectItem, string name)
        {
            List<string> itemNames = new List<string>();

            // 收集所有StyleItem
            List<StyleItem> allStyleItems = StylxTool.GetStyleItem(styleProjectItem);
            // 收集Name
            foreach (var styleItem in allStyleItems)
            {
                if (styleItem.Name == name)
                {
                    styleProjectItem.RemoveItem(styleItem);
                };
            }
        }


    }
}
