using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Extensions;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Dialogs;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using NPOI.OpenXmlFormats.Vml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using Geometry = ArcGIS.Core.Geometry.Geometry;
using MessageBox = System.Windows.MessageBox;

namespace CCTool.Scripts.LayerPross
{
    /// <summary>
    /// Interaction logic for GetPolygonExtent.xaml
    /// </summary>
    public partial class GetPolygonExtent : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public GetPolygonExtent()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "面要素计算四至";

        private async void btn_go_click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                bool isField = (bool)cb_origin.IsChecked;
                bool isPoint = (bool)cb_point.IsChecked;

                int digit = txtNum.Text.ToInt();

                // 获取参数listbox
                List<string> fieldNames = UITool.GetCheckboxStringFromListBox(listbox_field);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    Map map = MapView.Active.Map;
                    string gdbPath = Project.Current.DefaultGeodatabasePath;
                    // 获取图层
                    FeatureLayer featureLayer = MapView.Active.GetSelectedLayers().FirstOrDefault() as FeatureLayer;

                    string oid = featureLayer.Name.TargetIDFieldName();

                    // 如果选择的不是面要素或是无选择，则返回
                    if (featureLayer.ShapeType != esriGeometryType.esriGeometryPolygon || featureLayer == null)
                    {
                        MessageBox.Show("错误！请选择一个面要素！");
                        return;
                    }

                    pw.AddMessageStart("数据读取");

                    // 在原图层内标记四至点坐标
                    if (isField)
                    {
                        pw.AddMessageMiddle(20, "在原图层内标记四至点坐标");
                        // 添加字段
                        Arcpy.AddField(featureLayer.Name, "东至X", "DOUBLE");
                        Arcpy.AddField(featureLayer.Name, "东至Y", "DOUBLE");
                        Arcpy.AddField(featureLayer.Name, "南至X", "DOUBLE");
                        Arcpy.AddField(featureLayer.Name, "南至Y", "DOUBLE");
                        Arcpy.AddField(featureLayer.Name, "西至X", "DOUBLE");
                        Arcpy.AddField(featureLayer.Name, "西至Y", "DOUBLE");
                        Arcpy.AddField(featureLayer.Name, "北至X", "DOUBLE");
                        Arcpy.AddField(featureLayer.Name, "北至Y", "DOUBLE");

                        pw.AddMessageMiddle(10, "计算四至坐标");

                        using RowCursor rowCursor = featureLayer.Search();
                        while (rowCursor.MoveNext())
                        {
                            using Feature feature = rowCursor.Current as Feature;
                            // 标记一个初始坐标
                            double e_x = 0;
                            double n_y = 0;
                            double w_x = 100000000;
                            double s_y = 100000000;

                            Geometry geometry = feature.GetShape();
                            if (geometry is Polygon polygon)
                            {
                                // 找出四至点
                                foreach (var pt in polygon.Points)
                                {
                                    if (pt.X > e_x) { e_x = pt.X; }
                                    if (pt.Y > n_y) { n_y = pt.Y; }
                                    if (pt.X < w_x) { w_x = pt.X; }
                                    if (pt.Y < s_y) { s_y = pt.Y; }
                                }
                                // 标记四至点
                                foreach (var pt in polygon.Points)
                                {
                                    double xx = Math.Round(pt.X, digit);
                                    double yy = Math.Round(pt.Y, digit);

                                    if (pt.X == w_x)
                                    {
                                        feature["西至X"] = xx;
                                        feature["西至Y"] = yy;
                                    }
                                    if (pt.X == e_x)
                                    {
                                        feature["东至X"] = xx;
                                        feature["东至Y"] = yy;
                                    }
                                    if (pt.Y == n_y)
                                    {
                                        feature["北至X"] = xx;
                                        feature["北至Y"] = yy;
                                    }
                                    if (pt.Y == s_y)
                                    {
                                        feature["南至X"] = xx;
                                        feature["南至Y"] = yy;
                                    }
                                }
                            }
                            feature.Store();
                        }
                    }

                    // 导出四至点图层
                    if (isPoint)
                    {
                        pw.AddMessageMiddle(20, "导出四至点图层");

                        List<PointAtt> points = new List<PointAtt>();
                        
                        using RowCursor rowCursor = featureLayer.Search();
                        while (rowCursor.MoveNext())
                        {
                            using Feature feature = rowCursor.Current as Feature;

                            List<string> fieldValues = new List<string>();

                            // 保留字段
                            foreach (string fieldName in fieldNames)
                            {
                                string fieldValue = feature[fieldName]?.ToString();
                                if (fieldValue is null) { fieldValue = ""; }    // 避免空值
                                fieldValues.Add(fieldValue);
                            }

                            long oidValue = long.Parse(feature[oid].ToString());
                            // 标记一个初始坐标
                            double e_x = 0;
                            double n_y = 0;
                            double w_x = 100000000;
                            double s_y = 100000000;

                            Geometry geometry = feature.GetShape();
                            if (geometry is Polygon polygon)
                            {
                                // 找出四至点
                                foreach (var pt in polygon.Points)
                                {
                                    if (pt.X > e_x) { e_x = pt.X; }
                                    if (pt.Y > n_y) { n_y = pt.Y; }
                                    if (pt.X < w_x) { w_x = pt.X; }
                                    if (pt.Y < s_y) { s_y = pt.Y; }
                                }
                                // 标记四至点
                                foreach (var pt in polygon.Points)
                                {
                                    if (pt.X == w_x)
                                    {
                                        // 点信息
                                        PointAtt ptAtt = new PointAtt();
                                        ptAtt.X = pt.X;
                                        ptAtt.Y = pt.Y;
                                        ptAtt.Direction = "西";
                                        ptAtt.ID = oidValue;
                                        ptAtt.FieldValues = fieldValues;
                                        points.Add(ptAtt);
                                    }
                                    if (pt.X == e_x)
                                    {
                                        PointAtt ptAtt = new PointAtt();
                                        ptAtt.X = pt.X;
                                        ptAtt.Y = pt.Y;
                                        ptAtt.Direction = "东";
                                        ptAtt.ID = oidValue;
                                        ptAtt.FieldValues = fieldValues;
                                        points.Add(ptAtt);
                                    }
                                    if (pt.Y == n_y)
                                    {
                                        PointAtt ptAtt = new PointAtt();
                                        ptAtt.X = pt.X;
                                        ptAtt.Y = pt.Y;
                                        ptAtt.Direction = "北";
                                        ptAtt.ID = oidValue;
                                        ptAtt.FieldValues = fieldValues;
                                        points.Add(ptAtt);
                                    }
                                    if (pt.Y == s_y)
                                    {
                                        PointAtt ptAtt = new PointAtt();
                                        ptAtt.X = pt.X;
                                        ptAtt.Y = pt.Y;
                                        ptAtt.Direction = "南";
                                        ptAtt.ID = oidValue;
                                        ptAtt.FieldValues = fieldValues;
                                        points.Add(ptAtt);
                                    }

                                }
                            }
                        }

                        // 去重
                        // 按 Direction 分组后取每组第一个元素
                        List<PointAtt> distinctPoints = points
                            .GroupBy(p => new { p.Direction, p.ID })
                            .Select(g => g.First())
                            .ToList();

                        // 定义自定义方向顺序
                        var customOrder = new List<string> { "东", "南", "西", "北" };

                        // 多条件排序：先按ID升序，再按自定义方向顺序
                        List<PointAtt> sortedPoints = distinctPoints
                            .OrderBy(p => p.ID)                          // 第一排序条件：ID升序
                            .ThenBy(p => customOrder.IndexOf(p.Direction)) // 第二排序条件：自定义方向顺序
                            .ToList();


                        /// 创建点要素
                        // 创建一个ShapeDescription
                        // 获取坐标系
                        SpatialReference sr = featureLayer.GetSpatialReference();
                        // 导出的点要素名
                        string fcName = $"toPoint_{featureLayer.Name}";

                        // 判断一下是否存在目标要素，如果有的话，就删掉重建
                        bool isHaveTarget = gdbPath.IsHaveFeaturClass(fcName);
                        if (isHaveTarget)
                        {
                            Arcpy.Delect(@$"{gdbPath}\{fcName}");
                        }

                        var shapeDescription = new ShapeDescription(GeometryType.Point, sr)
                        {
                            HasM = false,
                            HasZ = false
                        };
                        // 定义4个字段
                        var polygonIndex = new ArcGIS.Core.Data.DDL.FieldDescription("原要素ID", FieldType.Integer);
                        var dh = new ArcGIS.Core.Data.DDL.FieldDescription("点号", FieldType.String);
                        var pointDir = new ArcGIS.Core.Data.DDL.FieldDescription("方位", FieldType.String);
                        var pointX = new ArcGIS.Core.Data.DDL.FieldDescription("X坐标", FieldType.Double);
                        var pointY = new ArcGIS.Core.Data.DDL.FieldDescription("Y坐标", FieldType.Double);

                        // 打开数据库gdb
                        using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath))))
                        {
                            // 收集字段列表
                            var fieldDescriptions = new List<ArcGIS.Core.Data.DDL.FieldDescription>()
                            {
                                polygonIndex,dh,pointDir,pointX, pointY
                            };

                            foreach (string fieldName in fieldNames)
                            {
                                fieldDescriptions.Add(new ArcGIS.Core.Data.DDL.FieldDescription(fieldName, FieldType.String));
                            }

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
                            /// 构建点要素
                            // 创建编辑操作对象
                            EditOperation editOperation = new EditOperation();
                            editOperation.Callback(context =>
                            {
                                // 获取要素定义
                                FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();

                                int dh = 1;
                                // 循环创建点
                                foreach (var pt in sortedPoints)
                                {
                                    // 创建RowBuffer
                                    using RowBuffer rowBuffer = featureClass.CreateRowBuffer();
                                    // 写入字段值
                                    rowBuffer["原要素ID"] = pt.ID;
                                    rowBuffer["点号"] = $"J{dh}";
                                    rowBuffer["方位"] = pt.Direction;
                                    rowBuffer["X坐标"] = Math.Round(pt.X, digit);
                                    rowBuffer["Y坐标"] = Math.Round(pt.Y, digit);

                                    for (int i = 0; i < fieldNames.Count; i++)
                                    {
                                        rowBuffer[fieldNames[i]] = pt.FieldValues[i];
                                    }

                                    // 创建点几何
                                    MapPointBuilderEx mapPointBuilderEx = new(new Coordinate2D(pt.X, pt.Y));
                                    // 给新添加的行设置形状
                                    rowBuffer[featureClassDefinition.GetShapeField()] = mapPointBuilderEx.ToGeometry();

                                    // 在表中创建新行
                                    using Feature feature = featureClass.CreateRow(rowBuffer);
                                    context.Invalidate(feature);      // 标记行为无效状态

                                    // 点号更新
                                    if (dh == 4)
                                    {
                                        dh = 1;
                                    }
                                    else
                                    {
                                        dh++;
                                    }
                                }

                            }, featureClass);

                            // 执行编辑操作
                            editOperation.Execute();
                            // 加载结果图层
                            MapCtlTool.AddLayerToMap(@$"{gdbPath}\{fcName}");
                        }

                        // 保存
                        Project.Current.SaveEditsAsync();

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

        class PointAtt
        {
            public long ID { get; set; }
            public string Direction { get; set; }
            public double X { get; set; }
            public double Y { get; set; }

            public List<string> FieldValues { get; set; }

        }

        private async void form_Load(object sender, RoutedEventArgs e)
        {
            // 获取图层
            FeatureLayer featureLayer = await QueuedTask.Run(() =>
            {
                return MapView.Active.GetSelectedLayers().FirstOrDefault() as FeatureLayer;
            });

            // 生成字段列表
            if (featureLayer is not null)
            {
                UITool.AddTextFieldsToListBox(listbox_field, featureLayer);
            }
        }

        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_field);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_field);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/148614668";
            UITool.Link2Web(url);
        }
    }
}
