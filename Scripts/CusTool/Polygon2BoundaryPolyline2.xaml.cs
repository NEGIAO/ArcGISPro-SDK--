using ActiproSoftware.Windows.Shapes;
using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using MathNet.Numerics.LinearAlgebra.Factorization;
using NPOI.OpenXmlFormats.Dml.Diagram;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.PTG;
using System;
using System.Collections.Generic;
using System.IO;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using Polygon = ArcGIS.Core.Geometry.Polygon;
using Polyline = ArcGIS.Core.Geometry.Polyline;
using SpatialReference = ArcGIS.Core.Geometry.SpatialReference;

namespace CCTool.Scripts.CusTool
{
    /// <summary>
    /// Interaction logic for Polygon2BoundaryPolyline2.xaml
    /// </summary>
    public partial class Polygon2BoundaryPolyline2 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public Polygon2BoundaryPolyline2()
        {
            InitializeComponent();

            // 初始化combox
            combox_sr.Items.Add("CGCS2000");
            combox_sr.Items.Add("北京54");
            combox_sr.Items.Add("西安80");
            combox_sr.Items.Add("WGS1984");
            combox_sr.SelectedIndex = 0;

            // 初始化输出路径
            string defGDB = Project.Current.DefaultGeodatabasePath;
            string defPath = Project.Current.HomeFolderPath;

            textExcelPath.Text = $@"{defPath}\界线描述表.xlsx";
            textPolylinePath.Text = $@"{defGDB}\导出界址线";
            textPointPath.Text = $@"{defGDB}\导出界址点";
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "界线导出Excel";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string in_fc = combox_fc.ComboxText();
                string excelPath = textExcelPath.Text;
                string bjField = combox_field.ComboxText();

                string srName = combox_sr.Text;

                int wid = srName switch
                {
                    "CGCS2000" => 4490,
                    "北京54" => 4214,
                    "西安80" => 4610,
                    "WGS1984" => 4326,
                    _ => 4490,
                };

                // 保存线，点
                string polylinePath = textPolylinePath.Text;
                string pointPath = textPointPath.Text;

                // 判断参数是否选择完全
                if (in_fc == "" || excelPath == "" || bjField == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                await QueuedTask.Run(() =>
                {
                    // 判断输入要素是否有Z值
                    if (in_fc.IsHasZ())
                    {
                        MessageBox.Show("输入的要素有Z值，请先清除掉！");
                        return;
                    }

                });

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart($"获取目标图层属性");
                    // 获取目标图层属性
                    Dictionary<List<PLAtt>, string> plAttList = GetAtt(in_fc, wid, bjField);

                    pw.AddMessageMiddle(20, $"生成界线描述表");
                    // 生成界址线表
                    CreateExcel(excelPath, plAttList);

                    // 生成界址线
                    if (polylinePath != "")
                    {
                        pw.AddMessageMiddle(20, $"生成界址线");
                        CreatePolyline(polylinePath, in_fc, plAttList);
                    }

                    // 生成界址点
                    if (pointPath != "")
                    {
                        pw.AddMessageMiddle(20, $"生成界址点");
                        CreatePoint(pointPath, in_fc, plAttList);
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
            string url = "https://blog.csdn.net/xcc34452366/article/details/138913392";
            UITool.Link2Web(url);
        }

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field);
        }

        // 判断方向
        public static string JudgeDirection(double angle)
        {
            string d = "";
            if (angle >= 348.75 || angle < 11.25)
            {
                d = "北";
            }
            else if (angle >= 11.25 && angle < 33.75)
            {
                d = "北东北";
            }
            else if (angle >= 33.75 && angle < 56.25)
            {
                d = "东北";
            }
            else if (angle >= 56.25 && angle < 78.75)
            {
                d = "东东北";
            }

            /////////
            if (angle >= 78.75 && angle < 101.25)
            {
                d = "东";
            }
            else if (angle >= 101.25 && angle < 123.75)
            {
                d = "东东南";
            }
            else if (angle >= 123.75 && angle < 146.25)
            {
                d = "东南";
            }
            else if (angle >= 146.25 && angle < 168.75)
            {
                d = "南东南";
            }


            /////////
            if (angle >= 168.75 && angle < 191.25)
            {
                d = "南";
            }
            else if (angle >= 191.25 && angle < 213.75)
            {
                d = "南西南";
            }
            else if (angle >= 213.75 && angle < 236.25)
            {
                d = "西南";
            }
            else if (angle >= 236.25 && angle < 258.75)
            {
                d = "西西南";
            }


            /////////
            if (angle >= 258.75 && angle < 281.25)
            {
                d = "西";
            }
            else if (angle >= 281.25 && angle < 303.75)
            {
                d = "西西北";
            }
            else if (angle >= 303.75 && angle < 326.25)
            {
                d = "西北";
            }
            else if (angle >= 326.25 && angle < 348.75)
            {
                d = "北西北";
            }

            return d;
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.SaveDialogExcel();
        }

        private void openPolylineButton_Click(object sender, RoutedEventArgs e)
        {
            textPolylinePath.Text = UITool.SaveDialogFeatureClass();
        }

        private void openPointButton_Click(object sender, RoutedEventArgs e)
        {
            textPointPath.Text = UITool.SaveDialogFeatureClass();
        }

        // 获取目标属性
        private Dictionary<List<PLAtt>, string> GetAtt(string in_fc, int wid, string bjField)
        {
            // 获取目标FeatureLayer
            FeatureLayer featurelayer = in_fc.TargetFeatureLayer();

            SpatialReference sr = SpatialReferenceBuilder.CreateSpatialReference(wid);
            // 预设界址线属性
            Dictionary<List<PLAtt>, string> plAttList = new Dictionary<List<PLAtt>, string>();

            // 获取所有的界址线信息
            RowCursor cursor = featurelayer.TargetSelectCursor();
            while (cursor.MoveNext())
            {
                List<PLAtt> plAtts = new List<PLAtt>();
                using var feature = cursor.Current as Feature;

                // 获取要素的几何
                Polygon polygon = feature.GetShape() as Polygon;

                // 坐标系转成地理坐标
                Polygon projectPolygon = GeometryEngine.Instance.Project(polygon, sr) as Polygon;

                // 获取面要素的所有点，并按逆时针排序
                List<List<MapPoint>> mapPoints = projectPolygon.ReshotMapPointWise();
                // 原来的点集也留一个，用来算界线距离
                List<List<MapPoint>> originMapPoints = polygon.ReshotMapPointWise();

                for (int i = 0; i < mapPoints.Count; i++)
                {
                    // 当前环的线段号
                    int lineIndex = 1;
                    for (int j = 0; j < mapPoints[i].Count - 1; j++)
                    {
                        PLAtt plAtt = new PLAtt();
                        // 写入点集和标记字段
                        plAtt.MapPoints=new List<MapPoint> { originMapPoints[i][j], originMapPoints[i][j+1] };
                        // 写入点号
                        plAtt.PolylineIndex = lineIndex;

                        // 计算线段起始点的经纬度坐标
                        double x = mapPoints[i][j].X;
                        double y = mapPoints[i][j].Y;
                        double xx = mapPoints[i][j + 1].X;
                        double yy = mapPoints[i][j + 1].Y;

                        // 写入线段起始点的经纬度
                        plAtt.StartX = x;
                        plAtt.StartY = y;
                        plAtt.EndX = xx;
                        plAtt.EndY = yy;

                        plAtt.StartLat = x.ToDecimal();
                        plAtt.StartLng = y.ToDecimal();
                        plAtt.EndLat = xx.ToDecimal();
                        plAtt.EndLng = yy.ToDecimal();

                        // 计算线段起始点的平面投影坐标
                        double ox = originMapPoints[i][j].X;
                        double oy = originMapPoints[i][j].Y;
                        double oxx = originMapPoints[i][j + 1].X;
                        double oyy = originMapPoints[i][j + 1].Y;
                        // 计算当起终点的距离
                        double distance = Math.Round(Math.Sqrt(Math.Pow(oxx - ox, 2) + Math.Pow(oyy - oy, 2)), 2);
                        // 写入距离
                        plAtt.Distance = distance;

                        // 获取起终点的角度
                        List<double> xy1 = new List<double>() { mapPoints[i][j].X, mapPoints[i][j].Y };
                        List<double> xy2 = new List<double>() { mapPoints[i][j + 1].X, mapPoints[i][j + 1].Y };

                        double angle = BaseTool.CalculateAngleFromNorth(xy1, xy2);

                        // 写入角度
                        plAtt.Angle = $"{Math.Round(angle, 2)}°";

                        // 判断方向
                        string direction = JudgeDirection(angle);
                        // 写入方向
                        plAtt.Direction = $"{direction}";
                        // 加入集合
                        plAtts.Add(plAtt);

                        // 线段号加一
                        lineIndex++;
                    }
                }

                // 获取标记字段
                string bj = feature[bjField]?.ToString();

                // 加入到集合
                plAttList.Add(plAtts, bj);
            }

            return plAttList;
        }

        // 生成界址线表
        private static void CreateExcel(string excelPath, Dictionary<List<PLAtt>, string> plAttList)
        {
            // 复制Excel表
            DirTool.CopyResourceFile(@"CCTool.Data.Excel.界线描述表.xlsx", excelPath);
            // 写入Excel
            foreach (var plAttPairs in plAttList)
            {
                List<PLAtt> plAtts = plAttPairs.Key;
                string bj = plAttPairs.Value;
                // 复制一个sheet
                ExcelTool.CopySheet(excelPath, "sheet1", bj);

                // 打开工作薄
                Workbook wb = ExcelTool.OpenWorkbook(excelPath);
                // 打开工作表
                Worksheet worksheet = wb.Worksheets[bj];
                // 获取Cells
                Cells cells = worksheet.Cells;
                // 当前行
                int initRow = 2;
                for (int i = 0; i < plAtts.Count; i++)
                {
                    // 如果不是第一行，就复制一行
                    if (i != 0)
                    {
                        cells.CopyRow(cells, 2, initRow);
                    }
                    // 写入点号
                    cells[initRow, 0].Value = $"{plAtts[i].PolylineIndex}";
                    // 写入线段起始点的经纬度
                    cells[initRow, 1].Value = $"{plAtts[i].StartLat}";
                    cells[initRow, 2].Value = $"{plAtts[i].StartLng}";
                    cells[initRow, 3].Value = $"{plAtts[i].EndLat}";
                    cells[initRow, 4].Value = $"{plAtts[i].EndLng}";
                    // 写入距离
                    cells[initRow, 7].Value = $"{plAtts[i].Distance}";
                    // 写入角度
                    cells[initRow, 5].Value = $"{plAtts[i].Angle}";
                    // 写入方向
                    cells[initRow, 6].Value = $"{plAtts[i].Direction}";

                    // 当前行进一行
                    initRow++;
                }

                // 保存
                wb.Save(excelPath);
                wb.Dispose();
            }

            // 删除sheet1
            ExcelTool.DeleteSheet(excelPath, "sheet1");
        }

        // 生成界址线
        private static void CreatePolyline(string polylinePath, string in_fc, Dictionary<List<PLAtt>, string> plAttList)
        {
            string gdbPath = polylinePath[..(polylinePath.LastIndexOf(@".gdb") + 4)];
            string fcName = polylinePath[(polylinePath.LastIndexOf(@"\") + 1)..];
            // 如果已有要素，则清除
            bool isHaveTarget = gdbPath.IsHaveFeaturClass(fcName);

            if (isHaveTarget)
            {
                Arcpy.Delect(polylinePath);
            }
            /// 创建线要素
            // 创建一个ShapeDescription
            var sr2 = in_fc.GetFeatureClassAtt().SpatialReference;
            var shapeDescription = new ShapeDescription(GeometryType.Polyline, sr2)
            {
                HasM = false,
                HasZ = false
            };
            // 定义字段
            var bj = new ArcGIS.Core.Data.DDL.FieldDescription("标记", FieldType.String);
            var polylineIndex = new ArcGIS.Core.Data.DDL.FieldDescription("编号", FieldType.Integer);
            var startLat = new ArcGIS.Core.Data.DDL.FieldDescription("经度_起点", FieldType.String);
            var startLng = new ArcGIS.Core.Data.DDL.FieldDescription("纬度_起点", FieldType.String);
            var endLat = new ArcGIS.Core.Data.DDL.FieldDescription("经度_终点", FieldType.String);
            var endLng = new ArcGIS.Core.Data.DDL.FieldDescription("纬度_终点", FieldType.String);
            var angle = new ArcGIS.Core.Data.DDL.FieldDescription("方位角", FieldType.String);
            var direction = new ArcGIS.Core.Data.DDL.FieldDescription("方向", FieldType.String);
            var distance = new ArcGIS.Core.Data.DDL.FieldDescription("长度", FieldType.Double);

            // 打开数据库gdb
            using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath))))
            {
                // 收集字段列表
                var fieldDescriptions = new List<ArcGIS.Core.Data.DDL.FieldDescription>()
                            {
                                bj, polylineIndex, startLat, startLng, endLat, endLng, angle, direction, distance
                            };

                // 创建FeatureClassDescription
                var fcDescription = new FeatureClassDescription(fcName, fieldDescriptions, shapeDescription);
                // 创建SchemaBuilder
                SchemaBuilder schemaBuilder = new SchemaBuilder(gdb);
                // 将创建任务添加到DDL任务列表中
                schemaBuilder.Create(fcDescription);
                // 执行DDL
                bool success = schemaBuilder.Build();

                // 创建要素并添加到要素类中
                using FeatureClass featureClass = gdb.OpenDataset<FeatureClass>(fcName);
                /// 构建线要素
                // 创建编辑操作对象
                EditOperation editOperation = new EditOperation();
                editOperation.Callback(context =>
                {
                    // 获取要素定义
                    FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();
                    // 循环创建点
                    foreach (var plAttPairs in plAttList)
                    {
                        List<PLAtt> plAtts = plAttPairs.Key;
                        string bj = plAttPairs.Value;

                        foreach (var plAtt in plAtts)
                        {
                            // 创建RowBuffer
                            using RowBuffer rowBuffer = featureClass.CreateRowBuffer();
                            // 创建点集合
                            MapPoint initPt = plAtt.MapPoints[0];
                            MapPoint nextPt = plAtt.MapPoints[1];

                            List<MapPoint> points = new List<MapPoint>
                                        {
                                            initPt, nextPt
                                        };

                            // 写入字段值
                            rowBuffer["标记"] = bj;
                            rowBuffer["编号"] = plAtt.PolylineIndex;
                            rowBuffer["经度_起点"] = plAtt.StartLat;
                            rowBuffer["纬度_起点"] = plAtt.StartLng;
                            rowBuffer["经度_终点"] = plAtt.EndLat;
                            rowBuffer["纬度_终点"] = plAtt.EndLng;
                            rowBuffer["方位角"] = plAtt.Angle;
                            rowBuffer["方向"] = plAtt.Direction;
                            rowBuffer["长度"] = plAtt.Distance;

                            // 创建线几何
                            Polyline polylineWithAttrs = PolylineBuilderEx.CreatePolyline(points);

                            // 给新添加的行设置形状
                            rowBuffer[featureClassDefinition.GetShapeField()] = polylineWithAttrs;

                            // 在表中创建新行
                            using Feature feature = featureClass.CreateRow(rowBuffer);
                            context.Invalidate(feature);      // 标记行为无效状态
                        }

                    }

                }, featureClass);

                // 执行编辑操作
                editOperation.Execute();
                // 加载结果图层
                MapCtlTool.AddLayerToMap(polylinePath);
            }

            // 保存
            Project.Current.SaveEditsAsync();
        }

        // 生成界址线
        private static void CreatePoint(string pointPath, string in_fc, Dictionary<List<PLAtt>, string> plAttList)
        {
            string gdbPath = pointPath[..(pointPath.LastIndexOf(@".gdb") + 4)];
            string fcName = pointPath[(pointPath.LastIndexOf(@"\") + 1)..];
            // 如果已有要素，则清除
            bool isHaveTarget = gdbPath.IsHaveFeaturClass(fcName);

            if (isHaveTarget)
            {
                Arcpy.Delect(pointPath);
            }
            /// 创建线要素
            // 创建一个ShapeDescription
            var sr2 = in_fc.GetFeatureClassAtt().SpatialReference;
            var shapeDescription = new ShapeDescription(GeometryType.Point, sr2)
            {
                HasM = false,
                HasZ = false
            };
            // 定义字段
            var bj = new ArcGIS.Core.Data.DDL.FieldDescription("标记", FieldType.String);
            var polylineIndex = new ArcGIS.Core.Data.DDL.FieldDescription("编号", FieldType.Integer);
            var startLat = new ArcGIS.Core.Data.DDL.FieldDescription("经度", FieldType.String);
            var startLng = new ArcGIS.Core.Data.DDL.FieldDescription("纬度", FieldType.String);

            // 打开数据库gdb
            using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath))))
            {
                // 收集字段列表
                var fieldDescriptions = new List<ArcGIS.Core.Data.DDL.FieldDescription>()
                            {
                                bj, polylineIndex, startLat, startLng
                            };

                // 创建FeatureClassDescription
                var fcDescription = new FeatureClassDescription(fcName, fieldDescriptions, shapeDescription);
                // 创建SchemaBuilder
                SchemaBuilder schemaBuilder = new SchemaBuilder(gdb);
                // 将创建任务添加到DDL任务列表中
                schemaBuilder.Create(fcDescription);
                // 执行DDL
                bool success = schemaBuilder.Build();

                // 创建要素并添加到要素类中
                using FeatureClass featureClass = gdb.OpenDataset<FeatureClass>(fcName);
                /// 构建线要素
                // 创建编辑操作对象
                EditOperation editOperation = new EditOperation();
                editOperation.Callback(context =>
                {
                    // 获取要素定义
                    FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();
                    // 循环创建点
                    foreach (var plAttPairs in plAttList)
                    {
                        List<PLAtt> plAtts = plAttPairs.Key;
                        string bj = plAttPairs.Value;

                        foreach (var plAtt in plAtts)
                        {
                            // 创建RowBuffer
                            using RowBuffer rowBuffer = featureClass.CreateRowBuffer();
                            // 创建点集合
                            MapPoint initPt = plAtt.MapPoints[0];

                            // 写入字段值
                            rowBuffer["标记"] = bj;
                            rowBuffer["编号"] = plAtt.PolylineIndex;
                            rowBuffer["经度"] = plAtt.StartLat;
                            rowBuffer["纬度"] = plAtt.StartLng;

                            // 给新添加的行设置形状
                            rowBuffer[featureClassDefinition.GetShapeField()] = initPt;

                            // 在表中创建新行
                            using Feature feature = featureClass.CreateRow(rowBuffer);
                            context.Invalidate(feature);      // 标记行为无效状态
                        }
                    }

                }, featureClass);

                // 执行编辑操作
                editOperation.Execute();
                // 加载结果图层
                MapCtlTool.AddLayerToMap(pointPath);
            }

            // 保存
            Project.Current.SaveEditsAsync();
        }

    }
}


// 界址线属性
public class PLAtt
{
    public long PolylineIndex { get; set; }
    public string StartLat { get; set; }
    public string StartLng { get; set; }
    public string EndLat { get; set; }
    public string EndLng { get; set; }
    public string Angle { get; set; }
    public string Direction { get; set; }
    public double Distance { get; set; }
    public List<MapPoint> MapPoints { get; set; }
    public double StartX { get; set; }
    public double StartY { get; set; }
    public double EndX { get; set; }
    public double EndY { get; set; }
}