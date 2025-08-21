using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
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
using System.Windows.Shapes;
using Polyline = ArcGIS.Core.Geometry.Polyline;

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    /// <summary>
    /// Interaction logic for LineInterpolation.xaml
    /// </summary>
    public partial class LineInterpolation : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public LineInterpolation()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "计算线要素的插值点数据";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc, "Polyline");
        }

        private void combox_start_DropDown(object sender, EventArgs e)
        {
            UITool.AddFieldsToComboxPlus(combox_fc.ComboxText(), combox_start);
        }

        private void combox_end_DropDown(object sender, EventArgs e)
        {
            UITool.AddFieldsToComboxPlus(combox_fc.ComboxText(), combox_end);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/148235155";
            UITool.Link2Web(url);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc = combox_fc.ComboxText();

                string startField = combox_start.ComboxText();
                string endField = combox_end.ComboxText();

                double interval = textDistance.Text.ToDouble();

                bool isPoint = (bool)cb_point.IsChecked;   // 按折点插值


                // 判断参数是否选择完全
                if (fc == "" || startField == "" || endField == "" | interval <= 0)
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

                    string pointName = "插值点";
                    string resultField = "计算值";

                    if (isPoint) // 如果是按折点插值
                    {
                        pw.AddMessageMiddle(20, "在线段折点上插值");
                        CreatePointByBreak( fc, startField, endField, pointName, resultField);
                    }
                    else
                    {
                        pw.AddMessageMiddle(20, $"按{interval}米距离插值");
                        CreatePointByDistance(fc ,interval, startField, endField, pointName, resultField);
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

        // 处理河流左右岸线
        private void CreatePointByBreak(string lyName, string startField, string endField, string pointName, string resultField)
        {
            // 默认数据库位置
            var defGDB = Project.Current.DefaultGeodatabasePath;

            // 获取目标FeatureLayer
            FeatureLayer featurelayer = lyName.TargetFeatureLayer();

            // 创建点要素
            SpatialReference sr = featurelayer.GetSpatialReference();
            Arcpy.CreateFeatureclass(defGDB, pointName, "POINT", sr, "", sr.XYTolerance, sr.XYResolution, true);
            // 添加字段
            Arcpy.AddField(pointName, resultField, "DOUBLE");
            string parentOID = "原线段OID";
            Arcpy.AddField(pointName, parentOID, "LONG");

            // 创建编辑操作
            var editOperation = new EditOperation();
            editOperation.Name = "Create points along line";

            // 获取点要素定义
            List<(long oid, MapPoint pt, double gc)> points = new List<(long oid, MapPoint pt, double gc)>();

            FeatureClass featureClass = pointName.TargetFeatureClass();
            FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();

            // 分析线要素
            using var lineCursor = featurelayer.Search();
            while (lineCursor.MoveNext())
            {
                var lineFeature = lineCursor.Current as Feature;
                var line = lineFeature.GetShape() as Polyline;
                // OID
                long oid = lineFeature.GetObjectID();
                double minGC = (lineFeature[startField]?.ToString()).ToDouble();
                double maxGC = (lineFeature[endField]?.ToString()).ToDouble();

                // 总长
                double totalLength = GeometryEngine.Instance.Length(line);
                // 当前点的长度
                double initLength = 0;

                // 获取线要素的部件
                var collection = line.Parts.ToList().FirstOrDefault();

                // 折点的XY值
                double prevX = double.NaN;
                double prevY = double.NaN;

                // 每个环进行处理（第一个为外环，其它为内环）
                foreach (Segment segment in collection)
                {
                    MapPoint mapPoint = segment.StartPoint;     // 获取点

                    double x = mapPoint.X;
                    double y = mapPoint.Y;

                    // 计算当前点与上一个点的距离
                    if (!double.IsNaN(prevX) && !double.IsNaN(prevY))
                    {
                        double distance = Math.Sqrt(Math.Pow(x - prevX, 2) + Math.Pow(y - prevY, 2));
                        initLength += distance;
                    }
                    prevX = x;
                    prevY = y;

                    // 计算当前点的高程
                    double initGC = Math.Round((maxGC - minGC) / totalLength * initLength + minGC, 2);

                    points.Add((oid, mapPoint, initGC));
                }

                // 加入最末点
                points.Add((oid, line.Points.LastOrDefault(), maxGC));
            }

            /// 创建点
            foreach (var pointItem in points)
            {
                // 创建RowBuffer
                using RowBuffer rowBuffer = featureClass.CreateRowBuffer();
                // 写入高程字段
                rowBuffer[resultField] = pointItem.gc;
                rowBuffer[parentOID] = pointItem.oid;

                // 创建点几何
                MapPointBuilderEx mapPointBuilderEx = new(new Coordinate2D(pointItem.pt.X, pointItem.pt.Y));
                // 给新添加的行设置形状
                rowBuffer[featureClassDefinition.GetShapeField()] = mapPointBuilderEx.ToGeometry();

                // 在表中创建新行
                using Feature feature = featureClass.CreateRow(rowBuffer);
            }

            // 执行编辑操作
            editOperation.Execute();
            // 保存
            Project.Current.SaveEditsAsync();
        }

        // 处理河流中心线
        private void CreatePointByDistance(string lyName, double interval, string startField, string endField, string pointName, string resultField)
        {
            // 默认数据库位置
            var defGDB = Project.Current.DefaultGeodatabasePath;

            // 获取目标FeatureLayer
            FeatureLayer featurelayer = lyName.TargetFeatureLayer();

            // 创建点要素
            SpatialReference sr = featurelayer.GetSpatialReference();
            Arcpy.CreateFeatureclass(defGDB, pointName, "POINT", sr, "", sr.XYTolerance, sr.XYResolution, true);
            // 添加字段
            Arcpy.AddField(pointName, resultField, "DOUBLE");
            string parentOID = "原线段OID";
            Arcpy.AddField(pointName, parentOID, "LONG");

            // 创建编辑操作
            var editOperation = new EditOperation();
            editOperation.Name = "Create points along line";

            // 获取点要素定义
            List<(long oid, MapPoint pt, double gc)> points = new List<(long oid, MapPoint pt, double gc)>();

            FeatureClass featureClass = pointName.TargetFeatureClass();
            FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();

            // 分析线要素
            using var lineCursor = featurelayer.Search();
            while (lineCursor.MoveNext())
            {
                var lineFeature = lineCursor.Current as Feature;
                var line = lineFeature.GetShape() as Polyline;
                // OID
                long oid = lineFeature.GetObjectID();
                double minGC = (lineFeature[startField]?.ToString()).ToDouble();
                double maxGC = (lineFeature[endField]?.ToString()).ToDouble();

                // 总长
                double totalLength = GeometryEngine.Instance.Length(line);
                // 当前点的长度
                double initLength = -interval;
                // 按距离取点
                for (double distance = 0; distance <= totalLength; distance += interval)
                {
                    // 使用QueryPoint方法获取指定距离处的点
                    MapPoint pt = GeometryEngine.Instance.QueryPoint
                    (
                        line,
                        SegmentExtensionType.NoExtension,  // 不延伸线段
                        distance,
                        AsRatioOrLength.AsLength
                    );
                    // 计算当前点的高程
                    initLength += interval;
                    double initGC = Math.Round((maxGC - minGC) / totalLength * initLength + minGC, 2);

                    points.Add((oid, pt, initGC));
                }
                // 加入最末点
                points.Add((oid, line.Points.LastOrDefault(), maxGC));
            }

            /// 创建点
            foreach (var pointItem in points)
            {
                // 创建RowBuffer
                using RowBuffer rowBuffer = featureClass.CreateRowBuffer();
                // 写入高程字段
                rowBuffer[resultField] = pointItem.gc;
                rowBuffer[parentOID] = pointItem.oid;

                // 创建点几何
                MapPointBuilderEx mapPointBuilderEx = new(new Coordinate2D(pointItem.pt.X, pointItem.pt.Y));
                // 给新添加的行设置形状
                rowBuffer[featureClassDefinition.GetShapeField()] = mapPointBuilderEx.ToGeometry();

                // 在表中创建新行
                using Feature feature = featureClass.CreateRow(rowBuffer);
            }

            // 执行编辑操作
            editOperation.Execute();
            // 保存
            Project.Current.SaveEditsAsync();
        }

    }
}
