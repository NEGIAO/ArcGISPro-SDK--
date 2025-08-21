using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
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
using ArcGIS.Core.Geometry;
using System.Security.Policy;
using ActiproSoftware.Windows.Shapes;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using MathNet.Numerics.LinearAlgebra.Factorization;
using NPOI.Util;
using ArcGIS.Core.Internal.CIM;
using Segment = ArcGIS.Core.Geometry.Segment;
using ArcGIS.Desktop.Editing;
using NPOI.SS.Formula.PTG;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    /// <summary>
    /// Interaction logic for PolylineToPolygon.xaml
    /// </summary>
    public partial class PolylineToPolygon : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "PolylineToPolygon";

        public PolylineToPolygon()
        {
            InitializeComponent();

            text_distance.Text = BaseTool.ReadValueFromReg(toolSet, "minDistance");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "线转面(保留字段属性)";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc, "Polyline");
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                double minDistance = text_distance.Text.ToDouble();

                // 默认数据库位置
                var gdb_path = Project.Current.DefaultGeodatabasePath;
                // 工程默认文件夹位置
                string folder_path = Project.Current.HomeFolderPath;

                // 判断参数是否选择完全
                if (fc_path == "" || minDistance == 0)
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
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(fc_path);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10, $"创建空的面要素");

                    // 如果有目标要素，先清除
                    string fcName = $"{fc_path}_转面";
                    string outPath = $@"{gdb_path}\{fcName}";
                    bool isHaveTarget = gdb_path.IsHaveFeaturClass(fcName);
                    if (isHaveTarget) { Arcpy.Delect(outPath); }

                    // 复制要素
                    string fc_copy = $@"{Project.Current.DefaultGeodatabasePath}\fc_copy";
                    Arcpy.CopyFeatures(fc_path, fc_copy);

                    // 获取要素类的属性
                    FeatureClassAtt featureClassAtt = fc_copy.GetFeatureClassAtt();
                    /// 【创建面要素】
                    // 创建一个ShapeDescription
                    var shapeDescription = new ShapeDescription(GeometryType.Polygon, featureClassAtt.SpatialReference)
                    {
                        HasM = featureClassAtt.HasM,
                        HasZ = featureClassAtt.HasZ
                    };

                    // 打开默认数据库gdb
                    using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_path)));
                    // 创建FeatureClassDescription
                    var fcDescription = new FeatureClassDescription(fcName, featureClassAtt.FieldDescriptions, shapeDescription);
                    // 创建SchemaBuilder
                    SchemaBuilder schemaBuilder = new SchemaBuilder(gdb);
                    // 将创建任务添加到DDL任务列表中
                    schemaBuilder.Create(fcDescription);
                    // 执行DDL
                    schemaBuilder.Build();

                    pw.AddMessageMiddle(10, $"获取线要素的所有属性");

                    /// 【获取原数据属性表】
                    // 获取字段列表
                    List<string> fieldList = GisTool.GetFieldsNameFromTarget(fc_copy);
                    List<RowAtt> rowAtts = new List<RowAtt>();

                    FeatureClass featureClass = fc_copy.TargetFeatureClass();
                    // 遍历面要素类中的所有要素
                    RowCursor cursor = featureClass.Search();
                    while (cursor.MoveNext())
                    {
                        RowAtt rowAtt = new RowAtt();
                        Dictionary<string, string> fieldValues = new Dictionary<string, string>();
                        List<MapPoint> mapPoints = new List<MapPoint>();
                        using var feature = cursor.Current as Feature;
                        // 获取要素的几何
                        ArcGIS.Core.Geometry.Polyline polyline = feature.GetShape() as ArcGIS.Core.Geometry.Polyline;
                        // 获取字段值
                        foreach (string field in fieldList)
                        {
                            var target = feature[field];
                            string va = "";
                            if (target != null) { va = target.ToString(); }
                            fieldValues.Add(field, va);
                        }
                        rowAtt.FieldValues = fieldValues;
                        // 获取图形
                        // 获取线要素的部件（内外环）
                        var points = polyline.Points;
                        foreach (MapPoint mapPoint in points)
                        {
                            // 加入点集
                            mapPoints.Add(mapPoint);
                        }
                        rowAtt.Points = mapPoints;

                        // 判断是否符合闭合条件【首末点距离小于minDistance】
                        double distance = BaseTool.CalculateDistance(mapPoints[0], mapPoints[mapPoints.Count - 1]);
                        if (distance <= minDistance) { rowAtt.IsClosed = true; }
                        else { rowAtt.IsClosed = false; }

                        // 加入集合
                        rowAtts.Add(rowAtt);
                    }

                    pw.AddMessageMiddle(20, $"写入所有图斑");

                    /// 【生成面要素】
                    // 打开数据库
                    using (Geodatabase gdb2 = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_path))))
                    {
                        // 创建要素并添加到要素类中
                        using FeatureClass featureClass2 = gdb2.OpenDataset<FeatureClass>(fcName);
                        // 计数
                        long index = 1;
                        // 解析数据，创建面要素
                        foreach (var rowAtt in rowAtts)
                        {
                            // 计数标志
                            if (index % 100 == 0)
                            {
                                pw.AddMessageMiddle(0, $"累计图斑数量：{index}", Brushes.Gray);
                            }
                            // 如果不符合闭合条件，就跳过
                            if (!rowAtt.IsClosed) { continue; }

                            // 构建坐标点集合
                            var mapPoints_list = rowAtt.Points;
                            var vertices_list = new List<Coordinate2D>();
                            foreach (MapPoint mapPoints in mapPoints_list)
                            {
                                vertices_list.Add(new Coordinate2D(mapPoints.X, mapPoints.Y));
                            }

                            /// 构建面要素
                            // 创建编辑操作对象
                            EditOperation editOperation = new EditOperation();

                            // 获取要素定义
                            FeatureClassDefinition featureClassDefinition = featureClass2.GetDefinition();
                            // 创建RowBuffer
                            using RowBuffer rowBuffer = featureClass2.CreateRowBuffer();

                            // 写入字段值
                            Dictionary<string, string> fieldValues = rowAtt.FieldValues;
                            foreach (var fieldValue in fieldValues)
                            {
                                // 写入的时候，跳过空值
                                if (fieldValue.Value != "")
                                {
                                    rowBuffer[fieldValue.Key] = fieldValue.Value;
                                }
                            }

                            PolygonBuilderEx pb = new PolygonBuilderEx(vertices_list);

                            // 给新添加的行设置形状
                            rowBuffer[featureClassDefinition.GetShapeField()] = pb.ToGeometry();

                            // 在表中创建新行
                            using Feature feature = featureClass2.CreateRow(rowBuffer);

                            // 执行编辑操作
                            editOperation.Execute();

                            index ++;
                        }
                    }
                    // 保存编辑
                    Project.Current.SaveEditsAsync();
                    // 修复几何
                    Arcpy.RepairGeometry(outPath);

                    Arcpy.Delect(fc_copy);

                    // 加载
                    MapCtlTool.AddLayerToMap(outPath);

                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }


        private List<string> CheckData(string in_data)
        {
            List<string> result = new List<string>();


            return result;
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/141418991";
            UITool.Link2Web(url);
        }
    }
}

public class RowAtt
{
    public Dictionary<string, string> FieldValues { get; set; }
    public List<MapPoint> Points { get; set; }
    public bool IsClosed { get; set; }
}