using ActiproSoftware.Windows.Shapes;
using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using MathNet.Numerics.LinearAlgebra.Factorization;
using Microsoft.Office.Core;
using NPOI.OpenXmlFormats.Dml.Diagram;
using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.Formula.PTG;
using System;
using System.Collections.Generic;
using System.IO;
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
using Geometry = ArcGIS.Core.Geometry.Geometry;
using Polygon = ArcGIS.Core.Geometry.Polygon;
using Polyline = ArcGIS.Core.Geometry.Polyline;
using SpatialReference = ArcGIS.Core.Geometry.SpatialReference;

namespace CCTool.Scripts.CusTool4
{
    /// <summary>
    /// Interaction logic for ExportPointAndPolyline.xaml
    /// </summary>
    public partial class ExportPointAndPolyline : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "ExportPointAndPolyline";
        public ExportPointAndPolyline()
        {
            InitializeComponent();

            textPointPath.Text = BaseTool.ReadValueFromReg(toolSet, "outPoint");
            textPolylinePath.Text = BaseTool.ReadValueFromReg(toolSet, "outPolyline");

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "[确权登记]生成界址点、界址线";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }


        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc = combox_fc.ComboxText();
                string outPoint = textPointPath.Text;
                string outPolyline = textPolylinePath.Text;

                // 判断参数是否选择完全
                if (fc == "" || outPoint == "" || outPolyline == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                BaseTool.WriteValueToReg(toolSet, "outPoint", outPoint);
                BaseTool.WriteValueToReg(toolSet, "outPolyline", outPolyline);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(fc, outPoint, outPolyline);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    // 判断一下是否存在目标要素，如果有的话，就删掉重建
                    CheckTargetFC(outPoint);
                    CheckTargetFC(outPolyline);

                    // 获取目标FeatureLayer
                    FeatureLayer featurelayer = fc.TargetFeatureLayer();
                    // 获取坐标系
                    SpatialReference sr = featurelayer.GetSpatialReference();
                    // 确保要素类的几何类型是多边形
                    if (featurelayer.ShapeType != esriGeometryType.esriGeometryPolygon)
                    {
                        // 如果不是多边形类型，则输出错误信息并退出函数
                        MessageBox.Show("该要素类不是多边形类型。");
                        return;
                    }

                    pw.AddMessageMiddle(10, $"处理面要素，按西北角起始，顺时针重排");

                    Dictionary<List<List<MapPoint>>, string> mapPoints = new Dictionary<List<List<MapPoint>>, string>();
                    // 遍历面要素类中的所有要素
                    RowCursor cursor = featurelayer.TargetSelectCursor();
                    while (cursor.MoveNext())
                    {
                        using var feature = cursor.Current as Feature;
                        // 获取要素的几何
                        Polygon geometry = feature.GetShape() as Polygon;

                        if (geometry != null)
                        {
                            string feature_name = "";
                            // 获取面要素的所有折点【按西北角起始，顺时针重排】
                            mapPoints.Add(geometry.ReshotMapPoint(), feature_name);
                        }

                    }

                    pw.AddMessageMiddle(10, "【界址点】");
                    pw.AddMessageMiddle(0, "      创建点要素", Brushes.Gray);
                    // 创建点要素
                    CreatePoint(outPoint, sr, mapPoints);

                    pw.AddMessageMiddle(10, "      去重", Brushes.Gray);
                    // 去重
                    Arcpy.DeleteIdentical(outPoint);

                    pw.AddMessageMiddle(10, "      计算字段值", Brushes.Gray);
                    // 字段处理
                    CalulateFieldPoint(outPoint, fc, pw);

                    pw.AddMessageMiddle(10, "【界址线】");
                    pw.AddMessageMiddle(0, "      创建线要素", Brushes.Gray);
                    // 创建线要素
                    CreatePolyline(outPolyline, sr, mapPoints);

                    pw.AddMessageMiddle(10, "      去重", Brushes.Gray);
                    // 去重
                    Arcpy.DeleteIdentical(outPolyline);

                    pw.AddMessageMiddle(10, "      计算字段值", Brushes.Gray);
                    // 字段处理
                    CalulateFieldPolyline(outPolyline, fc);

                    CalulateFieldPolyline2(outPolyline, outPoint);

                    // 加载
                    MapCtlTool.AddLayerToMap(outPoint);
                    MapCtlTool.AddLayerToMap(outPolyline);
                });

                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        // 字段处理
        private void CalulateFieldPoint(string outPoint, string fc, ProcessWindow pw)
        {
            Arcpy.CalculateField(outPoint, "JZDH", "1000000+!OBJECTID!");
            Arcpy.CalculateField(outPoint, "BSM", "5000000+!OBJECTID!");

            string defGDB = Project.Current.DefaultGeodatabasePath;
            string fcCopy = $@"{defGDB}\fcCopy";
            string joinFeature = $@"{defGDB}\joinFeature";
            // 空间连接
            Arcpy.CopyFeatures(fc, fcCopy);
            Arcpy.AddField(fcCopy, "标记", "TEXT");
            Arcpy.CalculateField(fcCopy, "标记", "!DKBM!");
            Arcpy.DeleteField(fcCopy, "标记", "KEEP_FIELDS");
            // 比较复杂的空间连接，需要文字连接
            GPExecuteToolFlags executeFlags = Arcpy.SetGPFlag(false);

            string exp = $"BSM \"标识码\" true true false 4 Long 0 0,First,#,point,BSM,-1,-1;YSDM \"要素代码\" true true false 6 Text 0 0,First,#,point,YSDM,0,5;JZDH \"界址点号\" true true false 10 Text 0 0,First,#,point,JZDH,0,9;JZDLX \"界址点类型\" true true false 1 Text 0 0,First,#,point,JZDLX,0,254;JBLX \"界标类型\" true true false 1 Text 0 0,First,#,point,JBLX,0,254;DKBM \"地块代码\" true true false 254 Text 0 0,Join,\"/\",{fcCopy},标记,0,254;XZBZ \"X坐标值\" true true false 8 Double 0 0,First,#,point,XZBZ,-1,-1;YZBZ \"Y坐标值\" true true false 8 Double 0 0,First,#,point,YZBZ,-1,-1;Shape_Length \"Shape_Length\" false true true 8 Double 0 0,First,#,fcCopy,Shape_Length,-1,-1;Shape_Area \"Shape_Area\" false true true 8 Double 0 0,First,#,fcCopy,Shape_Area,-1,-1;标记 \"标记\" true true false 255 Text 0 0,First,#,fcCopy,标记,0,254";
            var par = Geoprocessing.MakeValueArray(outPoint, fcCopy, joinFeature, "JOIN_ONE_TO_ONE", "KEEP_ALL", exp, "INTERSECT", "", "");
            Geoprocessing.ExecuteToolAsync("analysis.SpatialJoin", par, null, null, null, executeFlags);

            Arcpy.DeleteField(joinFeature, new List<string>() { "Join_Count", "TARGET_FID", "标记"});
            Arcpy.CopyFeatures(joinFeature, outPoint);

            Arcpy.Delect(joinFeature);
            Arcpy.Delect(fcCopy);

        }

        // 字段处理
        private void CalulateFieldPolyline(string outPolyline, string fc)
        {
            Arcpy.CalculateField(outPolyline, "JZXH", "1000000+!OBJECTID!");
            Arcpy.CalculateField(outPolyline, "BSM", "2000000+!OBJECTID!");

            string defGDB = Project.Current.DefaultGeodatabasePath;
            string fcCopy = $@"{defGDB}\fcCopy";
            string joinFeature = $@"{defGDB}\joinFeature";
            // 空间连接
            Arcpy.CopyFeatures(fc, fcCopy);
            Arcpy.AddField(fcCopy, "编码", "TEXT");
            Arcpy.AddField(fcCopy, "姓名", "TEXT");
            Arcpy.CalculateField(fcCopy, "编码", "!DKBM!");
            Arcpy.CalculateField(fcCopy, "姓名", "!ZJRXM!");
            Arcpy.DeleteField(fcCopy, new List<string>() { "编码", "姓名" }, "KEEP_FIELDS");
            // 比较复杂的空间连接，需要文字连接
            GPExecuteToolFlags executeFlags = Arcpy.SetGPFlag(false);

            string exp = $"BSM \"标识码\" true true false 4 Long 0 0,First,#,C:\\Users\\Administrator\\Desktop\\导出\\控规.gdb\\outPolyline,BSM,-1,-1;YSDM \"要素代码\" true true false 6 Text 0 0,First,#,outPolyline,YSDM,0,5;JXXZ \"界线性质\" true true false 6 Text 0 0,First,#,outPolyline,JXXZ,0,5;JZXLB \"界址线类别\" true true false 2 Text 0 0,First,#,outPolyline,JZXLB,0,1;JZXWZ \"界址线位置\" true true false 1 Text 0 0,First,#,outPolyline,JZXWZ,0,254;JZXSM \"界址线说明\" true true false 254 Text 0 0,First,#,outPolyline,JZXSM,0,253;PLDWQLR \"毗邻地物权利人\" true true false 100 Text 0 0,Join,\"/\",C:\\Users\\Administrator\\Documents\\ArcGIS\\Projects\\Test\\Test.gdb\\fcCopy,姓名,0,254;PLDWZJR \"毗邻地物指界人\" true true false 100 Text 0 0,Join,\"/\",C:\\Users\\Administrator\\Documents\\ArcGIS\\Projects\\Test\\Test.gdb\\fcCopy,姓名,0,254;JZXH \"界址线号\" true true false 10 Text 0 0,First,#,outPolyline,JZXH,0,9;QJZDH \"起界址点号\" true true false 10 Text 0 0,First,#,outPolyline,QJZDH,0,9;ZJZDH \"止界址点号\" true true false 10 Text 0 0,First,#,outPolyline,ZJZDH,0,9;DKBM \"地块代码\" true true false 254 Text 0 0,Join,\"/\",C:\\Users\\Administrator\\Documents\\ArcGIS\\Projects\\Test\\Test.gdb\\fcCopy,编码,0,254;Shape_Length \"Shape_Length\" false true true 8 Double 0 0,First,#,outPolyline,Shape_Length,-1,-1;编码 \"编码\" true true false 255 Text 0 0,First,#,fcCopy,编码,0,254;姓名 \"姓名\" true true false 255 Text 0 0,First,#,fcCopy,姓名,0,254";
            var par = Geoprocessing.MakeValueArray(outPolyline, fcCopy, joinFeature, "JOIN_ONE_TO_ONE", "KEEP_ALL", exp, "INTERSECT", "", "");
            Geoprocessing.ExecuteToolAsync("analysis.SpatialJoin", par, null, null, null, executeFlags);

            Arcpy.DeleteField(joinFeature, new List<string>() { "Join_Count", "TARGET_FID", "编码", "姓名" });
            Arcpy.CopyFeatures(joinFeature, outPolyline);

            Arcpy.Delect(joinFeature);
            Arcpy.Delect(fcCopy);

        }

        // 字段处理
        private void CalulateFieldPolyline2(string outPolyline, string outPoint)
        {
            string defGDB = Project.Current.DefaultGeodatabasePath;
            string pointCopy = $@"{defGDB}\pointCopy";
            string joinFeature = $@"{defGDB}\joinFeature";
            // 空间连接
            Arcpy.CopyFeatures(outPoint, pointCopy);
            Arcpy.AddField(pointCopy, "标记", "TEXT");

            Arcpy.CalculateField(pointCopy, "标记", "!JZDH!");

            Arcpy.DeleteField(pointCopy, new List<string>() { "标记" }, "KEEP_FIELDS");
            // 比较复杂的空间连接，需要文字连接
            GPExecuteToolFlags executeFlags = Arcpy.SetGPFlag(false);

            string exp = $"BSM \"标识码\" true true false 4 Long 0 0,First,#,polyline,BSM,-1,-1;YSDM \"要素代码\" true true false 6 Text 0 0,First,#,polyline,YSDM,0,5;JXXZ \"界线性质\" true true false 6 Text 0 0,First,#,polyline,JXXZ,0,5;JZXLB \"界址线类别\" true true false 2 Text 0 0,First,#,polyline,JZXLB,0,1;JZXWZ \"界址线位置\" true true false 1 Text 0 0,First,#,polyline,JZXWZ,0,254;JZXSM \"界址线说明\" true true false 254 Text 0 0,First,#,polyline,JZXSM,0,253;PLDWQLR \"毗邻地物权利人\" true true false 100 Text 0 0,First,#,polyline,PLDWQLR,0,99;PLDWZJR \"毗邻地物指界人\" true true false 100 Text 0 0,First,#,polyline,PLDWZJR,0,99;JZXH \"界址线号\" true true false 10 Text 0 0,First,#,polyline,JZXH,0,9;QJZDH \"起界址点号\" true true false 10 Text 0 0,First,#,polyline,QJZDH,0,9;ZJZDH \"止界址点号\" true true false 10 Text 0 0,First,#,polyline,ZJZDH,0,9;DKBM \"地块代码\" true true false 254 Text 0 0,First,#,polyline,DKBM,0,253;Shape_Length \"Shape_Length\" false true true 8 Double 0 0,First,#,polyline,Shape_Length,-1,-1;标记 \"标记\" true true false 255 Text 0 0,Join,\",\",pointCopy,标记,0,254";
            var par = Geoprocessing.MakeValueArray(outPolyline, pointCopy, joinFeature, "JOIN_ONE_TO_ONE", "KEEP_ALL", exp, "INTERSECT", "", "");
            Geoprocessing.ExecuteToolAsync("analysis.SpatialJoin", par, null, null, null, executeFlags);

            // 计算字段
            Arcpy.CalculateField(joinFeature, "QJZDH", "!标记!.split(',')[0]");
            Arcpy.CalculateField(joinFeature, "ZJZDH", "!标记!.split(',')[1]");

            Arcpy.DeleteField(joinFeature, new List<string>() { "Join_Count", "TARGET_FID", "标记"});
            Arcpy.CopyFeatures(joinFeature, outPolyline);

            Arcpy.Delect(joinFeature);
            Arcpy.Delect(pointCopy);
        }

        // 创建点要素
        private void CreatePoint(string fcPath, SpatialReference sr, Dictionary<List<List<MapPoint>>, string> mapPoints)
        {
            string gdbPath = fcPath.TargetWorkSpace();
            string fcName = fcPath.TargetFcName();

            /// 创建点要素
            // 创建一个ShapeDescription
            var shapeDescription = new ShapeDescription(GeometryType.Point, sr)
            {
                HasM = false,
                HasZ = false
            };
            // 定义字段
            #region
            var bsm = new ArcGIS.Core.Data.DDL.FieldDescription("BSM", FieldType.Integer);
            bsm.AliasName = "标识码";
            bsm.Length = 10;

            var ysdm = new ArcGIS.Core.Data.DDL.FieldDescription("YSDM", FieldType.String);
            ysdm.AliasName = "要素代码";
            ysdm.Length = 6;

            var jzdh = new ArcGIS.Core.Data.DDL.FieldDescription("JZDH", FieldType.String);
            jzdh.AliasName = "界址点号";
            jzdh.Length = 10;

            var jzdlx = new ArcGIS.Core.Data.DDL.FieldDescription("JZDLX", FieldType.String);
            jzdlx.AliasName = "界址点类型";
            jzdlx.Length = 1;

            var jblx = new ArcGIS.Core.Data.DDL.FieldDescription("JBLX", FieldType.String);
            jblx.AliasName = "界标类型";
            jblx.Length = 1;

            var dkbm = new ArcGIS.Core.Data.DDL.FieldDescription("DKBM", FieldType.String);
            dkbm.AliasName = "地块代码";
            dkbm.Length = 254;

            var xzbz = new ArcGIS.Core.Data.DDL.FieldDescription("XZBZ", FieldType.Double);
            xzbz.AliasName = "X坐标值";
            xzbz.Length = 10;

            var yzbz = new ArcGIS.Core.Data.DDL.FieldDescription("YZBZ", FieldType.Double);
            yzbz.AliasName = "Y坐标值";
            yzbz.Length = 10;
            #endregion

            // 打开数据库gdb
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));
            // 收集字段列表
            var fieldDescriptions = new List<ArcGIS.Core.Data.DDL.FieldDescription>()
            {
                  bsm, ysdm, jzdh, jzdlx, jblx, dkbm, xzbz, yzbz,
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
            /// 构建点要素
            // 创建编辑操作对象
            EditOperation editOperation = new EditOperation();
            editOperation.Callback(context =>
            {
                // 获取要素定义
                FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();
                // 循环创建点
                foreach (var mp in mapPoints)
                {
                    var mpList = mp.Key;
                    for (int j = 0; j < mpList.Count; j++)
                    {
                        for (int k = 0; k < mpList[j].Count; k++)
                        {
                            // 创建RowBuffer
                            using RowBuffer rowBuffer = featureClass.CreateRowBuffer();
                            MapPoint pt = mpList[j][k];
                            // 写入字段值和标记代码
                            rowBuffer["YSDM"] = 211021;
                            rowBuffer["JZDLX"] = "3";
                            rowBuffer["JBLX"] = "9";

                            rowBuffer["XZBZ"] = Math.Round(pt.X, 3);
                            rowBuffer["YZBZ"] = Math.Round(pt.Y, 3);

                            // 坐标
                            Coordinate2D newCoordinate = new Coordinate2D(pt.X, pt.Y);
                            // 创建点几何
                            MapPointBuilderEx mapPointBuilderEx = new(new Coordinate2D(pt.X, pt.Y));
                            // 给新添加的行设置形状
                            rowBuffer[featureClassDefinition.GetShapeField()] = mapPointBuilderEx.ToGeometry();

                            // 在表中创建新行
                            using Feature feature = featureClass.CreateRow(rowBuffer);
                            context.Invalidate(feature);      // 标记行为无效状态
                        }
                    }
                }

            }, featureClass);

            // 执行编辑操作
            editOperation.Execute();
            // 保存
            Project.Current.SaveEditsAsync();
        }

        // 创建线要素
        private void CreatePolyline(string fcPath, SpatialReference sr, Dictionary<List<List<MapPoint>>, string> mapPoints)
        {
            string gdbPath = fcPath.TargetWorkSpace();
            string fcName = fcPath.TargetFcName();

            // 创建一个ShapeDescription
            var shapeDescription = new ShapeDescription(GeometryType.Polyline, sr)
            {
                HasM = false,
                HasZ = false
            };

            // 定义字段
            #region
            var bsm = new ArcGIS.Core.Data.DDL.FieldDescription("BSM", FieldType.Integer);
            bsm.AliasName = "标识码";
            bsm.Length = 10;

            var ysdm = new ArcGIS.Core.Data.DDL.FieldDescription("YSDM", FieldType.String);
            ysdm.AliasName = "要素代码";
            ysdm.Length = 6;

            var jxxz = new ArcGIS.Core.Data.DDL.FieldDescription("JXXZ", FieldType.String);
            jxxz.AliasName = "界线性质";
            jxxz.Length = 6;

            var jzxlb = new ArcGIS.Core.Data.DDL.FieldDescription("JZXLB", FieldType.String);
            jzxlb.AliasName = "界址线类别";
            jzxlb.Length = 2;

            var jzxwz = new ArcGIS.Core.Data.DDL.FieldDescription("JZXWZ", FieldType.String);
            jzxwz.AliasName = "界址线位置";
            jzxwz.Length = 1;

            var jzxsm = new ArcGIS.Core.Data.DDL.FieldDescription("JZXSM", FieldType.String);
            jzxsm.AliasName = "界址线说明";
            jzxsm.Length = 254;

            var pldwqlr = new ArcGIS.Core.Data.DDL.FieldDescription("PLDWQLR", FieldType.String);
            pldwqlr.AliasName = "毗邻地物权利人";
            pldwqlr.Length = 100;

            var pldwzjr = new ArcGIS.Core.Data.DDL.FieldDescription("PLDWZJR", FieldType.String);
            pldwzjr.AliasName = "毗邻地物指界人";
            pldwzjr.Length = 100;

            var jzxh = new ArcGIS.Core.Data.DDL.FieldDescription("JZXH", FieldType.String);
            jzxh.AliasName = "界址线号";
            jzxh.Length = 10;

            var qjzdh = new ArcGIS.Core.Data.DDL.FieldDescription("QJZDH", FieldType.String);
            qjzdh.AliasName = "起界址点号";
            qjzdh.Length = 10;

            var zjzdh = new ArcGIS.Core.Data.DDL.FieldDescription("ZJZDH", FieldType.String);
            zjzdh.AliasName = "止界址点号";
            zjzdh.Length = 10;

            var dkbm = new ArcGIS.Core.Data.DDL.FieldDescription("DKBM", FieldType.String);
            dkbm.AliasName = "地块代码";
            dkbm.Length = 254;

            #endregion

            // 打开数据库gdb
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));
            // 收集字段列表
            var fieldDescriptions = new List<ArcGIS.Core.Data.DDL.FieldDescription>()
            {
                bsm, ysdm, jxxz,jzxlb,jzxwz,jzxsm,pldwqlr,pldwzjr,jzxh,qjzdh,zjzdh,dkbm,
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
                foreach (var mp in mapPoints)
                {
                    var mpList = mp.Key;
                    for (int j = 0; j < mpList.Count; j++)
                    {
                        for (int k = 0; k < mpList[j].Count - 1; k++)
                        {
                            // 创建RowBuffer
                            using RowBuffer rowBuffer = featureClass.CreateRowBuffer();
                            // 创建点集合
                            MapPoint initPt = mpList[j][k];
                            MapPoint nextPt = mpList[j][k + 1];

                            List<MapPoint> points = new List<MapPoint>
                                    {
                                                initPt, nextPt
                                    };

                            // 写入字段值
                            rowBuffer["YSDM"] = 211031;
                            rowBuffer["JXXZ"] = "600001";
                            rowBuffer["JZXLB"] = "01";
                            rowBuffer["JZXWZ"] = "2";

                            // 创建线几何
                            Polyline polylineWithAttrs = PolylineBuilderEx.CreatePolyline(points);

                            // 给新添加的行设置形状
                            rowBuffer[featureClassDefinition.GetShapeField()] = polylineWithAttrs;

                            // 在表中创建新行
                            using Feature feature = featureClass.CreateRow(rowBuffer);
                            context.Invalidate(feature);      // 标记行为无效状态
                        }
                    }
                }

            }, featureClass);

            // 执行编辑操作
            editOperation.Execute();
            // 保存
            Project.Current.SaveEditsAsync();
        }


        private void CheckTargetFC(string outPoint)
        {
            // 获取目标数据库和点要素名
            string gdbPath = outPoint.TargetWorkSpace();
            string fcName = outPoint.TargetFcName();

            // 判断一下是否存在目标要素，如果有的话，就删掉重建
            bool isHaveTarget = gdbPath.IsHaveFeaturClass(fcName);

            if (isHaveTarget)
            {
                Arcpy.Delect(outPoint);
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147781798?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }


        private List<string> CheckData(string fc, string outPoint, string outPolyline)
        {
            List<string> result = new List<string>();

            // 判断输出路径是否为gdb
            string gdbResult = CheckTool.CheckGDBFeature(outPoint);
            if (gdbResult != "")
            {
                result.Add(gdbResult);
            }
            string gdbResult2 = CheckTool.CheckGDBFeature(outPolyline);
            if (gdbResult2 != "")
            {
                result.Add(gdbResult2);
            }

            // 判断输入要素是否有Z值
            string ZResult = CheckTool.CheckHasZ(fc);
            if (ZResult != "")
            {
                result.Add(ZResult);
            }

            string fieldResult = CheckTool.IsHaveFieldInTarget(fc, new List<string>() { "DKBM", "ZJRXM" });
            if (fieldResult != "")
            {
                result.Add(fieldResult);
            }


            return result;
        }

        private void openPointButton_Click(object sender, RoutedEventArgs e)
        {
            textPointPath.Text = UITool.SaveDialogFeatureClass();
        }

        private void openPolylineButton_Click(object sender, RoutedEventArgs e)
        {
            textPolylinePath.Text = UITool.SaveDialogFeatureClass();
        }
    }
}
