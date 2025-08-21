using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.SS.Formula.Functions;
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
using Polygon = ArcGIS.Core.Geometry.Polygon;
using Polyline = ArcGIS.Core.Geometry.Polyline;

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    /// <summary>
    /// Interaction logic for SearchShortLine.xaml
    /// </summary>
    public partial class SearchShortLine : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "SearchShortLine";

        public SearchShortLine()
        {
            InitializeComponent();

            txt_len.Text = BaseTool.ReadValueFromReg(toolSet, "len");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "查找弧线段";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc, "PP");
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string def_gdb = Project.Current.DefaultGeodatabasePath;    // 工程默认数据库

                string in_fc = combox_fc.ComboxText();
                int len = txt_len.Text.ToInt();   // 最小长度
                // 写入本地参数
                BaseTool.WriteValueToReg(toolSet, "len", len);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart($"获取目标FeatureLayer");
                    // 获取目标FeatureLayer
                    FeatureLayer featurelayer = in_fc.TargetFeatureLayer();
                    // 输出检查结果路径
                    string fcName = $@"{featurelayer.Name}_超短线";
                    string out_line = $@"{def_gdb}\{fcName}";

                    // 判断一下是否存在目标要素，如果有的话，就删掉重建
                    bool isHaveTarget = def_gdb.IsHaveFeaturClass(fcName);

                    if (isHaveTarget)
                    {
                        Arcpy.Delect(out_line);
                    }

                    // 获取坐标系
                    SpatialReference sr = featurelayer.GetSpatialReference();

                    List<Segment> lines = new();
                    // 遍历面要素类中的所有要素
                    RowCursor cursor = featurelayer.Search();
                    while (cursor.MoveNext())
                    {
                        using var feature = cursor.Current as Feature;
                        // 分线和面的情况
                        if (feature.GetShape() is Polygon polygon)
                        {
                            // 获取面要素的部件（内外环）
                            var parts = polygon.Parts.ToList();
                            // 获取超短线
                            lines.AddRange(GetShort(parts, len));
                        }
                        else if (feature.GetShape() is Polyline polyline)
                        {
                            // 获取面要素的部件（内外环）
                            var parts = polyline.Parts.ToList();
                            // 获取超短线
                            lines.AddRange(GetShort(parts, len));
                        }

                    }

                    pw.AddMessageMiddle(10, "创建线要素");


                    /// 创建线要素
                    // 创建一个ShapeDescription
                    var shapeDescription = new ShapeDescription(GeometryType.Polyline, sr)
                    {
                        HasM = false,
                        HasZ = false
                    };
                    // 定义4个字段
                    var pointIndex = new ArcGIS.Core.Data.DDL.FieldDescription("序号", FieldType.String);

                    // 打开数据库gdb
                    using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(def_gdb))))
                    {
                        // 收集字段列表
                        var fieldDescriptions = new List<ArcGIS.Core.Data.DDL.FieldDescription>()
                        {
                            pointIndex
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
                            foreach (var line in lines)
                            {
                                // 创建RowBuffer
                                using RowBuffer rowBuffer = featureClass.CreateRowBuffer();

                                rowBuffer["序号"] = 1;

                                // 创建线几何
                                Polyline polylineWithAttrs = PolylineBuilderEx.CreatePolyline(line);

                                // 给新添加的行设置形状
                                rowBuffer[featureClassDefinition.GetShapeField()] = polylineWithAttrs;

                                // 在表中创建新行
                                using Feature feature = featureClass.CreateRow(rowBuffer);
                                context.Invalidate(feature);      // 标记行为无效状态
                            }

                        }, featureClass);

                        // 执行编辑操作
                        editOperation.Execute();
                        // 加载结果图层
                        MapCtlTool.AddLayerToMap(out_line);
                    }

                    // 保存
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

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/146038325";
            UITool.Link2Web(url);
        }

        // 获取超短线
        private List<Segment> GetShort(List<ReadOnlySegmentCollection> parts, int len)
        {
            List<Segment> lines = new();

            foreach (ReadOnlySegmentCollection collection in parts)
            {
                List<MapPoint> points = new List<MapPoint>();
                // 每个环进行处理（第一个为外环，其它为内环）
                foreach (Segment segment in collection)
                {
                    if (segment.Length < len)
                    {
                        lines.Add(segment);
                    }
                }
            }

            return lines;
        }
    }
}
