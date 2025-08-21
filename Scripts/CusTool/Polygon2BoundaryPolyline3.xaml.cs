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
    /// Interaction logic for Polygon2BoundaryPolyline3.xaml
    /// </summary>
    public partial class Polygon2BoundaryPolyline3 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public Polygon2BoundaryPolyline3()
        {
            InitializeComponent();

            // 初始化combox
            combox_wise.Items.Add("顺时针");
            combox_wise.Items.Add("逆时针");
            combox_wise.SelectedIndex = 1;


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
            UITool.AddFeatureLayersToComboxPlus(combox_fc, "Polyline");
        }

        private void combox_point_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_point, "Point");
        }

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field);
        }

        private void combox_field_pbj_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field_pbj);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string in_line = combox_fc.ComboxText();
                string in_point = combox_point.ComboxText();
                string excelPath = textExcelPath.Text;
                string bjField = combox_field.ComboxText();
                string pbj = combox_field_pbj.ComboxText();
                string wise = combox_wise.Text;
                bool isWise = wise switch
                {
                    "顺时针" => true,
                    "逆时针" => false,
                    _ => false,
                };

                string defGDB = Project.Current.DefaultGeodatabasePath;
                string defFolder = Project.Current.HomeFolderPath;

                bool isClosed = (bool)cb_closeld.IsChecked;

                // 保存线，点
                string polylinePath = textPolylinePath.Text;
                string pointPath = textPointPath.Text;

                // 判断参数是否选择完全
                if (in_line == "" || excelPath == "" || bjField == "" || in_point == "")
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
                    pw.AddMessageStart($"按【{wise}】对边界点进行排序");

                    if (isClosed)
                    {
                        // 顺逆时针排序
                        in_line.SetPolylineWise(isWise);
                    }

                    pw.AddMessageMiddle(10, $"按点边界线分割");
                    string out_line = $@"{defGDB}\out_line";
                    string out_lineSort = $@"{defGDB}\out_lineSort";
                    // 按点边界线分割
                    Arcpy.SplitLineAtPoint(in_line, in_point, out_line);
                    // 排序
                    Arcpy.Sort(out_line, out_lineSort, $"{bjField} ASCENDING; ORIG_SEQ ASCENDING", "UR");

                    pw.AddMessageMiddle(10, $"获取目标图层属性");
                    // 获取目标图层属性
                    Dictionary<string, List<PLAtt2>> plAttList = GetAtt(out_lineSort, 4490, bjField, pbj);

                    pw.AddMessageMiddle(20, $"生成界线描述表");
                    // 生成界址线表
                    CreateExcel(excelPath, plAttList);

                    // 生成界址线
                    if (polylinePath != "")
                    {
                        pw.AddMessageMiddle(20, $"生成界址线");
                        CreatePolyline(polylinePath, out_lineSort, plAttList);
                    }

                    // 生成界址点
                    if (pointPath != "")
                    {
                        pw.AddMessageMiddle(20, $"生成界址点");
                        string point = $@"{defGDB}\point";
                        CreatePoint(point, out_lineSort, plAttList);
                        // 空间连接
                        Arcpy.SpatialJoin(in_point, point, pointPath, true);

                        Arcpy.Delect(point);
                    }

                    // 删除过程数据
                    Arcpy.Delect(out_line);
                    Arcpy.Delect(out_lineSort);
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
        private Dictionary<string, List<PLAtt2>> GetAtt(string in_fc, int wid, string bjField, string pbj)
        {
            // 获取目标FeatureClass
            FeatureClass featureClass = in_fc.TargetFeatureClass();

            SpatialReference sr_cgcs2000 = SpatialReferenceBuilder.CreateSpatialReference(wid);
            // 预设界址线属性
            Dictionary<string, List<PLAtt2>> plAttList = new Dictionary<string, List<PLAtt2>>();

            // 记录标记字段变化
            string bjValue = "";
            List<PLAtt2> plAtts = new List<PLAtt2>();
            // 获取所有的界址线信息
            RowCursor cursor = featureClass.Search();
            while (cursor.MoveNext())
            {
                using var feature = cursor.Current as Feature;
                string bj = feature[bjField]?.ToString();
                // 遇到新的一组线段，就清空
                if (bj != bjValue && plAtts.Count > 0)
                {
                    plAttList.Add(bjValue, plAtts);
                    plAtts = new List<PLAtt2>();
                }
                bjValue = bj;

                // 获取要素的几何
                Polyline polyline = feature.GetShape() as Polyline;

                // 坐标系转成投影坐标
                Polyline polyline_cgcs2000 = GeometryEngine.Instance.Project(polyline, sr_cgcs2000) as Polyline;

                PLAtt2 plAtt2 = new PLAtt2();

                // 全点集
                List<MapPoint> pts = polyline.Points.ToList();
                plAtt2.LineMapPoints = pts;
                // 首末点
                MapPoint startPoint = pts[0];
                MapPoint endPoint = pts[^1];    
                plAtt2.MapPoints = new List<MapPoint> { startPoint, endPoint };
                // 写入线号
                plAtt2.PolylineIndex = long.Parse(feature["ORIG_SEQ"]?.ToString());

                // 计算边界点编号
                string sortName = feature[pbj]?.ToString().GetWord("英文");
                int sortID = int.Parse(feature[pbj]?.ToString().GetWord("数字"));
                plAtt2.SortID = sortName + (sortID + plAtt2.PolylineIndex-1).ToString("D4");

                // 计算线段起始点的经纬度坐标
                double x = polyline_cgcs2000.Points[0].X;
                double y = polyline_cgcs2000.Points[0].Y;
                double xx = polyline_cgcs2000.Points[^1].X;
                double yy = polyline_cgcs2000.Points[^1].Y;

                // 写入线段起始点的经纬度
                plAtt2.StartX = x;
                plAtt2.StartY = y;
                plAtt2.EndX = xx;
                plAtt2.EndY = yy;

                plAtt2.StartLat = x.ToDecimal();
                plAtt2.StartLng = y.ToDecimal();
                plAtt2.EndLat = xx.ToDecimal();
                plAtt2.EndLng = yy.ToDecimal();

                // 计算线段起始点的平面投影坐标
                double ox = startPoint.X;
                double oy = startPoint.Y;
                double oxx = endPoint.X;
                double oyy = endPoint.Y;
                // 计算当起终点的距离
                double distance = Math.Round(Math.Sqrt(Math.Pow(oxx - ox, 2) + Math.Pow(oyy - oy, 2)), 2);
                // 写入距离
                plAtt2.Distance = distance;
                // 写入长度
                plAtt2.Length = polyline.Length;

                // 获取起终点的角度
                List<double> xy1 = new List<double>() { x, y };
                List<double> xy2 = new List<double>() { xx, yy };

                double angle = BaseTool.CalculateAngleFromNorth(xy1, xy2);

                // 写入角度
                plAtt2.Angle = $"{Math.Round(angle, 2)}°";

                // 判断方向
                string direction = JudgeDirection(angle);
                // 写入方向
                plAtt2.Direction = $"{direction}";
                // 加入集合
                plAtts.Add(plAtt2);
            }

            // 加入最后一个
            plAttList.Add(bjValue, plAtts);

            return plAttList;
        }


        // 生成界址线表
        private static void CreateExcel(string excelPath, Dictionary<string, List<PLAtt2>> plAttList)
        {
            // 复制Excel表
            DirTool.CopyResourceFile(@"CCTool.Data.Excel.界线描述表(念青东).xlsx", excelPath);

            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            Cells cells = sheet.Cells;

            // 当前行
            int initRow = 2;
            // 写入Excel
            foreach (var plAttPairs in plAttList)
            {
                List<PLAtt2> plAtts = plAttPairs.Value;
                string bj = plAttPairs.Key;

                for (int i = 0; i < plAtts.Count; i++)
                {
                    // 如果不是第一行，就复制一行
                    if (initRow != 2)
                    {
                        cells.CopyRow(cells, 2, initRow);
                    }
                    // 写入序号，边界线编号
                    cells[initRow, 0].Value = $"{i + 1}";
                    cells[initRow, 1].Value = $"{bj}";
                    // 起点边界点
                    string bjd = plAtts[i].SortID;
                    cells[initRow, 2].Value = $"{bjd}";
                    // 写入线段起始点的经纬度
                    cells[initRow, 3].Value = $"{plAtts[i].StartLat}";
                    cells[initRow, 4].Value = $"{plAtts[i].StartLng}";
                    cells[initRow, 5].Value = $"{plAtts[i].EndLat}";
                    cells[initRow, 6].Value = $"{plAtts[i].EndLng}";

                    // 写入角度
                    cells[initRow, 7].Value = $"{plAtts[i].Angle}";
                    // 写入方向
                    cells[initRow, 8].Value = $"{plAtts[i].Direction}";
                    // 写入长度
                    cells[initRow, 9].Value = $"{Math.Round(plAtts[i].Length, 2).RoundWithFill(2)}";
                    // 写入距离
                    cells[initRow, 10].Value = $"{Math.Round(plAtts[i].Distance, 2).RoundWithFill(2)}";

                    // 当前行进一行
                    initRow++;
                }
            }

            // 保存
            wb.Save(excelPath);
            wb.Dispose();
        }

        // 生成界址线
        private static void CreatePolyline(string polylinePath, string in_fc, Dictionary<string, List<PLAtt2>> plAttList)
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
            var bj = new ArcGIS.Core.Data.DDL.FieldDescription("边界线编号", FieldType.String);
            var polylineIndex = new ArcGIS.Core.Data.DDL.FieldDescription("起点边界点编号", FieldType.String);
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
                        List<PLAtt2> plAtts = plAttPairs.Value;
                        string bj = plAttPairs.Key;

                        foreach (var plAtt in plAtts)
                        {
                            // 创建RowBuffer
                            using RowBuffer rowBuffer = featureClass.CreateRowBuffer();

                            // 写入字段值
                            rowBuffer["边界线编号"] = bj;
                            rowBuffer["起点边界点编号"] = plAtt.SortID;
                            rowBuffer["经度_起点"] = plAtt.StartLat;
                            rowBuffer["纬度_起点"] = plAtt.StartLng;
                            rowBuffer["经度_终点"] = plAtt.EndLat;
                            rowBuffer["纬度_终点"] = plAtt.EndLng;
                            rowBuffer["方位角"] = plAtt.Angle;
                            rowBuffer["方向"] = plAtt.Direction;
                            rowBuffer["长度"] = plAtt.Distance;

                            // 创建线几何
                            Polyline polylineWithAttrs = PolylineBuilderEx.CreatePolyline(plAtt.LineMapPoints);

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
        private static void CreatePoint(string pointPath, string in_fc, Dictionary<string, List<PLAtt2>> plAttList)
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
            var bj = new ArcGIS.Core.Data.DDL.FieldDescription("边界线编号", FieldType.String);
            var polylineIndex = new ArcGIS.Core.Data.DDL.FieldDescription("起点边界点编号", FieldType.String);
            //var startLat = new ArcGIS.Core.Data.DDL.FieldDescription("经度", FieldType.String);
            //var startLng = new ArcGIS.Core.Data.DDL.FieldDescription("纬度", FieldType.String);

            // 打开数据库gdb
            using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath))))
            {
                // 收集字段列表
                var fieldDescriptions = new List<ArcGIS.Core.Data.DDL.FieldDescription>()
                            {
                                bj, polylineIndex, /*startLat, startLng*/
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
                        List<PLAtt2> plAtts = plAttPairs.Value;
                        string bj = plAttPairs.Key;

                        foreach (var plAtt in plAtts)
                        {
                            // 创建RowBuffer
                            using RowBuffer rowBuffer = featureClass.CreateRowBuffer();
                            // 创建点集合
                            MapPoint initPt = plAtt.MapPoints[0];

                            // 写入字段值
                            rowBuffer["边界线编号"] = bj;
                            rowBuffer["起点边界点编号"] = plAtt.SortID;
                            //rowBuffer["经度"] = plAtt.StartLat;
                            //rowBuffer["纬度"] = plAtt.StartLng;

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
                //// 加载结果图层
                //MapCtlTool.AddFeatureLayerToMap(pointPath);
            }

            // 保存
            Project.Current.SaveEditsAsync();
        }

        
    }
}


// 界址线属性
public class PLAtt2
{
    public long PolylineIndex { get; set; }
    public string StartLat { get; set; }
    public string StartLng { get; set; }
    public string EndLat { get; set; }
    public string EndLng { get; set; }
    public string Angle { get; set; }
    public string Direction { get; set; }
    public double Distance { get; set; }
    public double Length { get; set; }

    public double StartX { get; set; }
    public double StartY { get; set; }
    public double EndX { get; set; }
    public double EndY { get; set; }

    public string SortID { get; set; }

    public List<MapPoint> MapPoints { get; set; }
    public List<MapPoint> LineMapPoints { get; set; }
}