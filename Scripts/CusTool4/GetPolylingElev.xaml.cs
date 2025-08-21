using ActiproSoftware.Windows.Shapes;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.Util;
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
using System.Windows.Media.Media3D;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Polyline = ArcGIS.Core.Geometry.Polyline;
using Segment = ArcGIS.Core.Geometry.Segment;

namespace CCTool.Scripts.CusTool4
{
    /// <summary>
    /// Interaction logic for GetPolylingElev.xaml
    /// </summary>
    public partial class GetPolylingElev : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "GetPolylingElev";
        public GetPolylingElev()
        {
            InitializeComponent();

            textGDBPath.Text = BaseTool.ReadValueFromReg(toolSet, "gdbPath");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "计算节点高程(龙)";

        private void combox_left_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_left, "Polyline");
        }

        private void combox_center_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_center, "Polyline");
        }

        private void combox_right_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_right, "Polyline");
        }

        private void combox_start_DropDown(object sender, EventArgs e)
        {
            UITool.AddFieldsToComboxPlus(combox_left.ComboxText(), combox_start);
        }

        private void combox_end_DropDown(object sender, EventArgs e)
        {
            UITool.AddFieldsToComboxPlus(combox_left.ComboxText(), combox_end);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/148232580";
            UITool.Link2Web(url);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string left = combox_left.ComboxText();
                string center = combox_center.ComboxText();
                string right = combox_right.ComboxText();

                string startField = combox_start.ComboxText();
                string endField = combox_end.ComboxText();

                double interval = textDistance.Text.ToDouble();

                string gdbPath = textGDBPath.Text;


                // 判断参数是否选择完全
                if (left == "" || center == "" || right == "" || startField == "" || endField == "" || gdbPath == "" || interval <= 0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "gdbPath", gdbPath);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(center, right, startField, endField);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(20, "处理河流左岸线");
                    CreateLeftRigntPoint(gdbPath, left, startField, endField, "左岸线取点");

                    pw.AddMessageMiddle(20, "处理河流右岸线");
                    CreateLeftRigntPoint(gdbPath, right, startField, endField, "右岸线取点");

                    pw.AddMessageMiddle(20, "处理河流中心线");
                    CreateCenterPoint(gdbPath, center, interval, startField, endField);


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
        private void CreateLeftRigntPoint(string gdbPath, string lyName, string startField, string endField, string pointName)
        {
            string dmh = "DMH";

            // 默认数据库位置
            var defGDB = Project.Current.DefaultGeodatabasePath;

            string pointPath = $@"{defGDB}\{pointName}";
            // 获取目标FeatureLayer
            FeatureLayer featurelayer = lyName.TargetFeatureLayer();

            // 创建点要素
            Arcpy.CreateFeatureclass(defGDB, pointName, "POINT", featurelayer.GetSpatialReference());
            // 添加字段
            Arcpy.AddField(pointPath, "X", "DOUBLE");
            Arcpy.AddField(pointPath, "Y", "DOUBLE");
            Arcpy.AddField(pointPath, "Z", "DOUBLE");
            Arcpy.AddField(pointPath, dmh, "TEXT");


            // 创建编辑操作
            var editOperation = new EditOperation();
            editOperation.Name = "Create points along line";

            // 获取点要素定义
            List<(string mc, MapPoint pt, double gc)> points = new List<(string mc, MapPoint pt, double gc)>();

            FeatureClass featureClass = pointPath.TargetFeatureClass();
            FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();

            // 分析线要素
            using var lineCursor = featurelayer.Search();
            while (lineCursor.MoveNext())
            {
                var lineFeature = lineCursor.Current as Feature;
                var line = lineFeature.GetShape() as Polyline;
                // OID
                string mc = lineFeature[dmh]?.ToString();
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

                    points.Add((mc, mapPoint, initGC));
                }

                // 加入最末点
                points.Add((mc, line.Points.LastOrDefault(), maxGC));
            }

            /// 创建点
            foreach (var pointItem in points)
            {
                // 创建RowBuffer
                using RowBuffer rowBuffer = featureClass.CreateRowBuffer();
                // 写入高程字段
                rowBuffer["X"] = pointItem.pt.X;
                rowBuffer["Y"] = pointItem.pt.Y;
                rowBuffer["Z"] = pointItem.gc;
                rowBuffer[dmh] = pointItem.mc;

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

            // 导出至gdb
            Arcpy.CreateFileGDB(gdbPath, pointName);
            string gdb = $@"{gdbPath}\{pointName}.gdb";
            Arcpy.SplitByAttributes(pointPath, gdb, dmh);
        }

        // 处理河流中心线
        private void CreateCenterPoint(string gdbPath, string lyName, double interval, string startField, string endField)
        {
            string pointName = "中心线取点";
            string dmh = "DMH";

            // 默认数据库位置
            var defGDB = Project.Current.DefaultGeodatabasePath;
            string pointPath = $@"{defGDB}\{pointName}";
            // 获取目标FeatureLayer
            FeatureLayer featurelayer = lyName.TargetFeatureLayer();

            // 创建点要素
            Arcpy.CreateFeatureclass(defGDB, pointName, "POINT", featurelayer.GetSpatialReference());
            // 添加字段
            Arcpy.AddField(pointPath, "X", "DOUBLE");
            Arcpy.AddField(pointPath, "Y", "DOUBLE");
            Arcpy.AddField(pointPath, "Z", "DOUBLE");
            Arcpy.AddField(pointPath, dmh, "TEXT");


            // 创建编辑操作
            var editOperation = new EditOperation();
            editOperation.Name = "Create points along line";

            // 获取点要素定义
            List<(string mc, MapPoint pt, double gc)> points = new List<(string mc, MapPoint pt, double gc)>();

            FeatureClass featureClass = pointPath.TargetFeatureClass();
            FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();

            // 分析线要素
            using var lineCursor = featurelayer.Search();
            while (lineCursor.MoveNext())
            {
                var lineFeature = lineCursor.Current as Feature;
                var line = lineFeature.GetShape() as Polyline;
                // OID
                string mc = lineFeature[dmh]?.ToString();
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

                    points.Add((mc, pt, initGC));
                }
                // 加入最末点
                points.Add((mc, line.Points.LastOrDefault(), maxGC));
            }

            /// 创建点
            foreach (var pointItem in points)
            {
                // 创建RowBuffer
                using RowBuffer rowBuffer = featureClass.CreateRowBuffer();
                // 写入高程字段
                rowBuffer["X"] = pointItem.pt.X;
                rowBuffer["Y"] = pointItem.pt.Y;
                rowBuffer["Z"] = pointItem.gc;
                rowBuffer[dmh] = pointItem.mc;

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

            // 导出至gdb
            Arcpy.CreateFileGDB(gdbPath, pointName);
            string gdb = $@"{gdbPath}\{pointName}.gdb";
            Arcpy.SplitByAttributes(pointPath, gdb, dmh);
        }

        private List<string> CheckData(string center, string right, string startField, string endField)
        {
            List<string> result = new List<string>();
            // 检查字段值是否存在
            List<string> fields = new List<string>() { startField, endField };
            string result_value = CheckTool.IsHaveFieldInLayer(center, fields);
            if (result_value != "")
            {
                result.Add(result_value);
            }

            string result_value2 = CheckTool.IsHaveFieldInLayer(right, fields);
            if (result_value2 != "")
            {
                result.Add(result_value2);
            }

            return result;
        }

        private void openGDBButton_Click(object sender, RoutedEventArgs e)
        {
            textGDBPath.Text = UITool.OpenDialogFolder();
        }
    }
}
