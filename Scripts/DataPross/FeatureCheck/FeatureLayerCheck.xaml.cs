using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.SS.Formula.Functions;
using Org.BouncyCastle.Crypto;
using System;
using System.Collections.Generic;
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
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Geometry = ArcGIS.Core.Geometry.Geometry;
using MessageBox = System.Windows.Forms.MessageBox;
using Polygon = ArcGIS.Core.Geometry.Polygon;
using Row = ArcGIS.Core.Data.Row;
using Table = ArcGIS.Core.Data.Table;

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    /// <summary>
    /// Interaction logic for FeatureLayerCheck.xaml
    /// </summary>
    public partial class FeatureLayerCheck : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public FeatureLayerCheck()
        {
            InitializeComponent();

            // 初始化参数
            string setName = "PolygonCheckSet";
            txt_miniArea.Text = BaseTool.ReadValueFromReg(setName, "miniArea");
            txt_miniAngle.Text = BaseTool.ReadValueFromReg(setName, "miniAngle");

            string isGeo = BaseTool.ReadValueFromReg(setName, "isGeo");
            string isPart = BaseTool.ReadValueFromReg(setName, "isPart");
            string isHole = BaseTool.ReadValueFromReg(setName, "isHole");
            string isMiniArea = BaseTool.ReadValueFromReg(setName, "isMiniArea");

            string isArc = BaseTool.ReadValueFromReg(setName, "isArc");
            string isAcute = BaseTool.ReadValueFromReg(setName, "isAcute");

            cb_geo.IsChecked = isGeo == "True";
            cb_part.IsChecked = isPart == "True";
            cb_hole.IsChecked = isHole == "True";
            cb_miniArea.IsChecked = isMiniArea == "True";

            cb_arc.IsChecked = isArc == "True";
            cb_acute.IsChecked = isAcute == "True";
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "面图层简单检查";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc, "Polygon");
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                _ = double.TryParse(txt_miniArea.Text, out double miniArea);
                _ = double.TryParse(txt_miniAngle.Text, out double miniAngle);

                // 选项
                bool isGeo = (bool)cb_geo.IsChecked;
                bool isPart = (bool)cb_part.IsChecked;
                bool isHole = (bool)cb_hole.IsChecked;
                bool isMiniArea = (bool)cb_miniArea.IsChecked;

                bool isArc = (bool)cb_arc.IsChecked;
                bool isAcute = (bool)cb_acute.IsChecked;

                // 存储参数
                string setName = "PolygonCheckSet";
                BaseTool.WriteValueToReg(setName, "miniArea", miniArea);
                BaseTool.WriteValueToReg(setName, "miniAngle", miniAngle);

                BaseTool.WriteValueToReg(setName, "isGeo", isGeo);
                BaseTool.WriteValueToReg(setName, "isPart", isPart);
                BaseTool.WriteValueToReg(setName, "isHole", isHole);
                BaseTool.WriteValueToReg(setName, "isMiniArea", isMiniArea);

                BaseTool.WriteValueToReg(setName, "isArc", isArc);
                BaseTool.WriteValueToReg(setName, "isAcute", isAcute);


                // 判断参数是否选择完全
                if (fc_path == "")
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
                    // 检查几何错误
                    if (isGeo)
                    {
                        pw.AddMessageMiddle(10, "【检查几何错误】", Brushes.Blue);
                        string errs = CheckGeo(fc_path);
                        if (errs!="")
                        {
                            pw.AddMessageMiddle(0, errs, Brushes.Red);
                        }
                    }

                    // 检查多部件
                    if (isPart)
                    {
                        pw.AddMessageMiddle(10, "【检查多部件】", Brushes.Blue);
                        string errs = CheckPart(fc_path);
                        if (errs != "")
                        {
                            pw.AddMessageMiddle(0, errs, Brushes.Red);
                        }
                    }

                    // 检查内部空洞
                    if (isHole)
                    {
                        pw.AddMessageMiddle(10, "【检查内部空洞】", Brushes.Blue);
                        string errs = CheckHole(fc_path);
                        if (errs != "")
                        {
                            pw.AddMessageMiddle(0, errs, Brushes.Red);
                        }
                    }

                    // 检查碎图斑
                    if (isMiniArea)
                    {
                        pw.AddMessageMiddle(10, "【检查碎图斑】", Brushes.Blue);
                        string errs = CheckMiniArea(fc_path, miniArea);
                        if (errs != "")
                        {
                            pw.AddMessageMiddle(0, errs, Brushes.Red);
                        }
                    }

                    // 检查弧线段
                    if (isArc)
                    {
                        pw.AddMessageMiddle(10, "【检查弧线段】", Brushes.Blue);
                        string errs = CheckArc(fc_path);
                        if (errs != "")
                        {
                            pw.AddMessageMiddle(0, errs, Brushes.Red);
                        }
                    }

                    // 检查尖锐角
                    if (isAcute)
                    {
                        pw.AddMessageMiddle(10, "【检查尖锐角】", Brushes.Blue);
                        string errs = CheckAcute(fc_path, miniAngle);
                        if (errs != "")
                        {
                            pw.AddMessageMiddle(0, errs, Brushes.Red);
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


        // 检查几何错误
        public static string CheckGeo(string fc_path)
        {
            string result = "";
            // 默认数据库位置
            string out_table = $@"{Project.Current.DefaultGeodatabasePath}\out_table";
            Arcpy.CheckGeometry(fc_path, out_table);

            // 获取ID字段
            string IDField = fc_path.TargetIDFieldName();

            using Table table = out_table.TargetTable();
            using RowCursor rowCursor = table.Search();
            while (rowCursor.MoveNext())
            {
                using Row row = rowCursor.Current;
                // OID值
                string OID = row["FEATURE_ID"].ToString();
                // PROBLEM值
                string problem = row["PROBLEM"].ToString();

                result += $"({IDField}:{OID})：{problem}。\r";

            }

            return result;
        }

        // 检查多部件
        public static string CheckPart(string fc_path)
        {
            string result = "";

            // 获取ID字段
            string IDField = fc_path.TargetIDFieldName();

            using FeatureClass featureClass = fc_path.TargetFeatureClass();
            using RowCursor rowCursor = featureClass.Search();
            while (rowCursor.MoveNext())
            {
                using Feature feature = rowCursor.Current as Feature;
                // OID值
                string OID = feature[IDField].ToString();

                Polygon polygon = feature.GetShape() as Polygon;
                // 获取面要素的所有点（内外环）
                List<List<MapPoint>> mapPoints = polygon.MapPointsFromPolygon();

                List<List<MapPoint>> newPoints = new List<List<MapPoint>>();
                // 挑选合适的
                foreach (List<MapPoint> mapPoint in mapPoints)
                {
                    // 判断顺逆时针
                    bool isClockwise = mapPoint.IsColckwise();
                    if (isClockwise)
                    {
                        newPoints.Add(mapPoint);
                    }
                }

                if (newPoints.Count > 1)
                {
                    result += $"({IDField}:{feature[IDField]})：部件数为【{newPoints.Count}】。\r";
                }

            }

            return result;
        }

        // 检查内部空洞
        public static string CheckHole(string fc_path)
        {
            string result = "";

            // 获取ID字段
            string IDField = fc_path.TargetIDFieldName();

            using FeatureClass featureClass = fc_path.TargetFeatureClass();
            using RowCursor rowCursor = featureClass.Search();
            while (rowCursor.MoveNext())
            {
                using Feature feature = rowCursor.Current as Feature;
                // OID值
                string OID = feature[IDField].ToString();

                Polygon polygon = feature.GetShape() as Polygon;
                // 获取面要素的所有点（内外环）
                List<List<MapPoint>> mapPoints = polygon.MapPointsFromPolygon();

                // 挑选合适的
                foreach (List<MapPoint> mapPoint in mapPoints)
                {
                    // 判断顺逆时针
                    bool isClockwise = mapPoint.IsColckwise();
                    if (!isClockwise)
                    {
                        result += $"({IDField}:{feature[IDField]})：存在内部空洞。\r";
                        break;
                    }
                }

            }

            return result;
        }

        // 检查碎图斑
        public static string CheckMiniArea(string fc_path, double miniArea)
        {
            string result = "";

            // 获取ID字段
            string IDField = fc_path.TargetIDFieldName();

            using FeatureClass featureClass = fc_path.TargetFeatureClass();
            using RowCursor rowCursor = featureClass.Search();
            while (rowCursor.MoveNext())
            {
                using Feature feature = rowCursor.Current as Feature;
                // OID值
                string OID = feature[IDField].ToString();

                Polygon polygon = feature.GetShape() as Polygon;

                if (polygon.Area < miniArea)
                {
                    result += $"({IDField}:{feature[IDField]})：面积为{polygon.Area}平方米。\r";
                }
            }

            return result;
        }

        // 检查弧线段
        public static string CheckArc(string fc_path)
        {
            string result = "";

            // 获取ID字段
            string IDField = fc_path.TargetIDFieldName();

            using FeatureClass featureClass = fc_path.TargetFeatureClass();
            using RowCursor rowCursor = featureClass.Search();
            while (rowCursor.MoveNext())
            {
                using Feature feature = rowCursor.Current as Feature;
                // OID值
                string OID = feature[IDField].ToString();

                Polygon polygon = feature.GetShape() as Polygon;
                // 检查弧线段
                var parts = polygon.Parts.ToList();
                foreach (ReadOnlySegmentCollection collection in parts)
                {
                    List<MapPoint> points = new List<MapPoint>();
                    // 每个环进行处理（第一个为外环，其它为内环）
                    foreach (Segment segment in collection)
                    {
                        var lineType = segment.SegmentType;
                        if (lineType == SegmentType.EllipticArc)
                        {
                            result += $"({IDField}:{feature[IDField]})：存在弧线段。\r";
                            break;
                        }
                    }
                }
            }

            return result;
        }

        // 检查尖锐角
        public static string CheckAcute(string fc_path, double miniArc)
        {
            string result = "";

            // 获取ID字段
            string IDField = fc_path.TargetIDFieldName();

            using FeatureClass featureClass = fc_path.TargetFeatureClass();
            using RowCursor rowCursor = featureClass.Search();
            while (rowCursor.MoveNext())
            {
                using Feature feature = rowCursor.Current as Feature;
                // OID值
                string OID = feature[IDField].ToString();

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

                        var angle1 = Math.Atan2(p2.Y - p1.Y, p2.X - p1.X);
                        var angle2 = Math.Atan2(p3.Y - p2.Y, p3.X - p2.X);

                        var angle = Math.Abs((angle2 - angle1) * 180 / Math.PI);

                        // 确保角度在0到180度之间
                        if (angle > 180)
                            angle = 360 - angle;

                        // 符合条件的，加入到集合
                        if (Math.Abs(180 - angle) < miniArc)
                        {
                            result += $"({IDField}:{feature[IDField]})：存在尖锐角。\r";
                            break;
                        }
                    }
                }
            }

            return result;
        }


        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/144576906";
            UITool.Link2Web(url);
        }

    }
}
