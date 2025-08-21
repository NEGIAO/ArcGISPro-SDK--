using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Editing.Attributes;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using ArcGIS.Desktop.Core;
using ArcGIS.Core.Data.DDL;
using FieldDescription = ArcGIS.Core.Data.DDL.FieldDescription;
using ArcGIS.Desktop.Framework.Dialogs;
using ArcGIS.Core.CIM;
using Aspose.Cells.Drawing;
using CCTool.Scripts.ToolManagers.Extensions;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.ToolManagers.Managers;
using Aspose.Cells.Charts;
using Geometry = ArcGIS.Core.Geometry.Geometry;

namespace CCTool.Scripts.ToolManagers
{
    public class GisTool
    {
        // 清除GDB下的所有要素
        public static void ClearGDBItem(string gdbPath)
        {
            // 打开gdb
            using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath))))
            {
                // 获取要素类和独立表
                var fcDefinitions = gdb.GetDefinitions<FeatureClassDefinition>();
                var tableDefinitions = gdb.GetDefinitions<TableDefinition>();

                // 删除空要素和表
                if (fcDefinitions.Count != 0)
                {
                    foreach (var fcDefinition in fcDefinitions)
                    {
                        FeatureClass featureClass = gdb.OpenDataset<FeatureClass>(fcDefinition.GetName());
                        // 删除
                        Arcpy.Delect(featureClass);
                    }
                }
                if (tableDefinitions.Count != 0)
                {
                    foreach (var tableDefinition in tableDefinitions)
                    {
                        Table table = gdb.OpenDataset<Table>(tableDefinition.GetName());
                        // 删除
                        Arcpy.Delect(table);
                    }
                }
            };
        }

        // 获取面要素的最小图斑面积
        public static double GetPolygonMinArea(string lyName)
        {
            double minArea = 100000000;

            // 获取原始图层的要素类
            FeatureClass featureClass = lyName.TargetFeatureClass();

            // 获取原始图层的要素游标
            using (RowCursor cursor = featureClass.Search())
            {
                // 遍历源图层的要素
                while (cursor.MoveNext())
                {
                    using Feature feature = (Feature)cursor.Current;
                    // 获取源要素的几何
                    ArcGIS.Core.Geometry.Geometry geometry = feature.GetShape();
                    // 要素面积
                    double originArea = (geometry as ArcGIS.Core.Geometry.Polygon).Area;

                    if (originArea < minArea)
                    {
                        minArea = originArea;
                    }
                }
            }

            return minArea;
        }

        // 属性映射
        public static void AttributeMapper(string inData, string mapField, string valueField, string mapExcel)
        {
            // 从Excel中获取dict
            Dictionary<string, string> mapDict = ExcelTool.GetDictFromExcel(mapExcel);

            // 获取原始图层的要素类
            Table table = inData.TargetTable();

            // 获取原始图层的要素游标
            using RowCursor rowCursor = table.Search();
            // 遍历源图层的要素
            while (rowCursor.MoveNext())
            {
                using Row row = rowCursor.Current;
                // 获取值
                var map = row[mapField];

                if (map is not null)
                {
                    string mapValue = map.ToString();
                    if (mapDict.ContainsKey(mapValue))
                    {
                        // 更新
                        row[valueField] = mapDict[mapValue];
                        row.Store();
                    }
                }
            }
        }

        // 属性映射
        public static void AttributeMapper(string inData, string mapField, string valueField, Dictionary<string, string> mapDict)
        {
            // 获取原始图层的要素类
            Table table = inData.TargetTable();

            // 获取原始图层的要素游标
            using RowCursor rowCursor = table.Search();
            // 遍历源图层的要素
            while (rowCursor.MoveNext())
            {
                using Row row = rowCursor.Current;
                // 获取值
                var map = row[mapField];

                if (map is not null)
                {
                    string mapValue = map.ToString();
                    if (mapDict.ContainsKey(mapValue))
                    {
                        // 更新
                        row[valueField] = mapDict[mapValue];
                        row.Store();
                    }
                }
            }
        }

        // 复制一个新图层，并自动命名
        public static string CopyCosFeatureLayer(FeatureLayer inLayer, string markStr = "copy")
        {
            // 复制的目标路径
            string targetName = GetCosFeatureClassName(inLayer.Name, markStr);

            // 当前工程的默认数据库，用来存放
            string defGDB = Project.Current.DefaultGeodatabasePath;

            string targetPath = $@"{defGDB}\{targetName}";

            // 复制要素
            Arcpy.CopyFeatures(inLayer, targetPath, true);

            return targetPath;
        }

        // 查找GDB中是否有要素类，如果有，则自动重命名
        private static string GetCosFeatureClassName(string inData, string markStr)
        {
            // 当前工程的默认数据库，用来存放
            string defGDB = Project.Current.DefaultGeodatabasePath;

            // 获取要素名
            string targetName = inData.TargetFcName();
            // 如果是原始名，就加后缀
            if (!targetName.Contains($"_{markStr}"))
            {
                targetName += $"_{markStr}";
            }

            // 判断是否有目标要素
            bool isHaveFc = defGDB.IsHaveFeaturClass(targetName);
            // 如果有目标要素，获取要素名，并给新要素重命名
            if (isHaveFc)
            {
                // 获取后缀
                string str = targetName[(targetName.LastIndexOf("_") + 1)..];
                // 转成数字
                int index = str.ToInt();
                // 如果不是数字，就直接加数字后缀即可
                if (index == 0)
                {
                    targetName += "_1";
                }
                // 如果是数字，就增加序号
                else
                {
                    index += 1;
                    targetName = targetName[..(targetName.LastIndexOf("_") + 1)] + index.ToString();
                }

                // 再循环验证
                targetName = GetCosFeatureClassName($@"{defGDB}\{targetName}", markStr);
            }

            return targetName;

        }


        // 合并要素【按OID】
        public static void MergeFeatures(FeatureLayer featureLayer, List<long> objectIDs)
        {
            var mergeFeatures = new EditOperation();

            // 创建【inspector】实例
            var inspector = new Inspector();
            // 加载
            inspector.Load(featureLayer, objectIDs[0]);
            // 执行
            mergeFeatures.Merge(featureLayer, objectIDs, inspector);
            mergeFeatures.Execute();
            // 保存
            Project.Current.SaveEditsAsync();
        }

        // 合并要素【按OID】
        public static void MergeFeatures(FeatureLayer featureLayer, List<List<long>> objectIDs)
        {
            foreach (var item in objectIDs)
            {
                var mergeFeatures = new EditOperation();
                // 创建【inspector】实例
                var inspector = new Inspector();
                // 加载
                inspector.Load(featureLayer, objectIDs[0]);
                // 执行
                mergeFeatures.Merge(featureLayer, item, inspector);

                mergeFeatures.Execute();

            }
            // 保存
            Project.Current.SaveEditsAsync();
        }

        // 删除要素【按OID】
        public static void DelectFeature(FeatureLayer featureLayer, long oid)
        {
            var deleteFeatures = new EditOperation();

            // 删除要素中的某一行
            deleteFeatures.Delete(featureLayer, oid);
            deleteFeatures.Execute();
            // 保存
            Project.Current.SaveEditsAsync();
        }

        // 删除多个要素【按OID】
        public static void DelectFeatures(FeatureLayer featureLayer, List<long> oids)
        {
            var deleteFeatures = new EditOperation();
            // 删除要素中的某一行
            deleteFeatures.Delete(featureLayer, oids);
            deleteFeatures.Execute();
            // 保存
            Project.Current.SaveEditsAsync();
        }

        // 修改要素别名
        public static void AlterAliasName(string gdbPath, string fcName, string aliasName)
        {
            // 打开数据库gdb
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));
            // 获取FeatureClassDefinition和FeatureClassDescription
            FeatureClassDefinition fcDefinition = gdb.GetDefinition<FeatureClassDefinition>(fcName);
            FeatureClassDescription fcDescription = new FeatureClassDescription(fcDefinition);

            // 更改别名
            fcDescription.AliasName = aliasName;

            // 创建SchemaBuilder
            SchemaBuilder schemaBuilder = new SchemaBuilder(gdb);
            // 将创建任务添加到DDL任务列表中
            schemaBuilder.Modify(fcDescription);
            // 执行DDL
            schemaBuilder.Build();
        }

        // 修改独立表别名
        public static void AlterTableAliasName(string gdbPath, string tbName, string aliasName)
        {
            // 打开数据库gdb
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));

            TableDefinition tbDefinition = gdb.GetDefinition<TableDefinition>(tbName);
            TableDescription tbDescription = new TableDescription(tbDefinition);

            // 更改别名
            tbDescription.AliasName = aliasName;

            // 创建SchemaBuilder
            SchemaBuilder schemaBuilder = new SchemaBuilder(gdb);
            // 将创建任务添加到DDL任务列表中
            schemaBuilder.Modify(tbDescription);
            // 执行DDL
            schemaBuilder.Build();
        }

        // 通过MapPoints创建点要素
        public static void CreatePointFromMapPoint(List<MapPoint> mapPoints, SpatialReference sr, string gdbPath, string fcName)
        {
            /// 创建点要素
            // 创建一个ShapeDescription
            var shapeDescription = new ShapeDescription(GeometryType.Point, sr)
            {
                HasM = false,
                HasZ = false
            };
            // 定义4个字段
            var pointX = new ArcGIS.Core.Data.DDL.FieldDescription("x坐标", FieldType.Double);
            var pointY = new ArcGIS.Core.Data.DDL.FieldDescription("y坐标", FieldType.Double);

            // 打开数据库gdb
            using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath))))
            {
                // 收集字段列表
                var fieldDescriptions = new List<ArcGIS.Core.Data.DDL.FieldDescription>()
                {
                      pointX, pointY
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
                using (FeatureClass featureClass = gdb.OpenDataset<FeatureClass>(fcName))
                {
                    /// 构建点要素
                    // 创建编辑操作对象
                    EditOperation editOperation = new EditOperation();
                    editOperation.Callback(context =>
                    {
                        // 获取要素定义
                        FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();
                        // 循环创建点
                        for (int i = 0; i < mapPoints.Count; i++)
                        {
                            // 创建RowBuffer
                            using RowBuffer rowBuffer = featureClass.CreateRowBuffer();
                            MapPoint pt = mapPoints[i];
                            // 写入字段值
                            rowBuffer["x坐标"] = pt.X;
                            rowBuffer["y坐标"] = pt.Y;

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
                    }, featureClass);

                    // 执行编辑操作
                    editOperation.Execute();
                }
            }

            // 保存
            Project.Current.SaveEditsAsync();
        }


        // 添加字段
        public static void AddField(string targetPath, string fieldName, FieldType fieldType = FieldType.String, int filedLength = 255, string aliasName = "")
        {
            // 获取参数
            string gdbPath = targetPath.TargetWorkSpace();       // 数据库路径
            string fcName = targetPath.TargetFcName();     // 要素名称

            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));

            // 获取待修改要素的【FeatureClassDefinition】
            FeatureClassDefinition featureClassDefinition = gdb.GetDefinition<FeatureClassDefinition>(fcName);
            // 获取【FeatureClassDescription】
            FeatureClassDescription featureClassDescription = new FeatureClassDescription(featureClassDefinition);

            // 别名
            if (aliasName == "") { aliasName = fieldName; }

            // 定义需要添加的字段
            FieldDescription description = new FieldDescription(fieldName, fieldType)
            {
                AliasName = aliasName,
                Length = filedLength,
            };

            // 获取所有字段名
            List<string> fields = new List<string>();
            foreach (var item in featureClassDescription.FieldDescriptions)
            {
                fields.Add(item.Name);
            }

            // 如果要添加的字段不重名，就将新字段添加到【FieldDescription】列表中
            List<FieldDescription> modifiedFieldDescriptions = new List<FieldDescription>(featureClassDescription.FieldDescriptions);
            if (!fields.Contains(fieldName))
            {
                modifiedFieldDescriptions.Add(description);
            }

            // 使用新添加的字段创建一个【FeatureClassDescription】
            FeatureClassDescription modifiedFeatureClassDescription = new FeatureClassDescription(featureClassDescription.Name, modifiedFieldDescriptions, featureClassDescription.ShapeDescription);

            SchemaBuilder schemaBuilder = new SchemaBuilder(gdb);

            // 更新要素
            schemaBuilder.Modify(modifiedFeatureClassDescription);
            schemaBuilder.Build();
        }

        // 通过Lyrx文件应用符号系统【唯一值】
        public static void ApplySymbol(string lyName, string fieldName, string lyrxPath)
        {
            // 获取FeatureLayer
            FeatureLayer featureLayer = lyName.TargetFeatureLayer();
            // 获取Lyrx图层
            LayerDocument lyrFile = new LayerDocument(lyrxPath);
            CIMLayerDocument cimLyrDoc = lyrFile.GetCIMLayerDocument();
            // 以唯一值方式获取CIMFeatureLayer
            CIMUniqueValueRenderer uvr = ((CIMFeatureLayer)cimLyrDoc.LayerDefinitions[0]).Renderer as CIMUniqueValueRenderer;
            // 修改一个参照字段
            uvr.Fields = new string[] { fieldName };
            // 应用渲染器
            featureLayer.SetRenderer(uvr);
        }

        // 通过Lyrx文件应用符号系统【唯一值】
        public static void ApplySymbol(FeatureLayer featureLayer, string fieldName, string lyrxPath)
        {
            // 获取Lyrx图层
            LayerDocument lyrFile = new LayerDocument(lyrxPath);
            CIMLayerDocument cimLyrDoc = lyrFile.GetCIMLayerDocument();
            // 以唯一值方式获取CIMFeatureLayer
            CIMUniqueValueRenderer uvr = ((CIMFeatureLayer)cimLyrDoc.LayerDefinitions[0]).Renderer as CIMUniqueValueRenderer;
            // 修改一个参照字段
            uvr.Fields = new string[] { fieldName };
            // 应用渲染器
            featureLayer.SetRenderer(uvr);
        }

        // 通过Lyrx文件应用符号系统【单一符号】
        public static void ApplySymbol(FeatureLayer featureLayer, string lyrxPath)
        {
            // 获取Lyrx图层
            LayerDocument lyrFile = new LayerDocument(lyrxPath);
            CIMLayerDocument cimLyrDoc = lyrFile.GetCIMLayerDocument();
            // 以唯一值方式获取CIMFeatureLayer
            CIMSimpleRenderer sr = ((CIMFeatureLayer)cimLyrDoc.LayerDefinitions[0]).Renderer as CIMSimpleRenderer;
            // 应用渲染器
            featureLayer.SetRenderer(sr);
        }

        // 符号系统删除计数为0的值
        public static void Delete0uvClass(FeatureLayer featureLayer)
        {
            // 获取CIMUniqueValueRenderer
            CIMUniqueValueRenderer uvr = featureLayer.GetRenderer() as CIMUniqueValueRenderer;

            // 获取映射字段
            string mapFieldName = uvr.Fields.FirstOrDefault();
            // 获取字段值映射表
            List<string> listFieldValues = featureLayer.GetFieldValues(mapFieldName);

            CIMUniqueValueClass[] uvClasses = uvr.Groups[0].Classes;

            // 删除计数值为0的行
            uvr.Groups[0].Classes = uvClasses.Where(x => listFieldValues.Contains(x.Values[0].FieldValues[0].ToString())).ToArray();

            // 应用渲染器
            featureLayer.SetRenderer(uvr);
        }

        // 删除字段，只针对gdb数据库【Delete：删除，Keep：保留】
        public static void DeleteField(string in_table, string fieldName, string model = "Delete")
        {
            string gdbPath = in_table.TargetWorkSpace();
            string fcName = in_table.TargetFcName();

            using Geodatabase geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));

            TableDefinition tableDefinition = geodatabase.GetDefinition<TableDefinition>(fcName);
            // 创建SchemaBuilder
            SchemaBuilder schemaBuilder = new SchemaBuilder(geodatabase);

            TableDescription tableDescription = new TableDescription(tableDefinition);

            // 从列表中找到要删除的字段
            List<FieldDescription> fieldsToBeRetained = new List<FieldDescription>() { };
            if (model == "Keep")
            {
                FieldDescription taxFieldToBeRetained = new FieldDescription(tableDefinition.GetFields().First(f => f.Name.Equals(fieldName)));
                fieldsToBeRetained.Add(taxFieldToBeRetained);
            }
            else
            {
                IReadOnlyList<Field> fields = tableDefinition.GetFields();
                foreach (Field field in fields)
                {
                    if (field.FieldType == FieldType.Geometry || field.IsEditable == false)     // 不可编辑的字段
                    {
                        continue;
                    }
                    if (field.Name != fieldName)
                    {
                        fieldsToBeRetained.Add(new FieldDescription(field));
                    }
                }
            }

            // 更新TableDescription
            TableDescription modifiedTableDescription = new TableDescription(tableDescription.Name, fieldsToBeRetained);
            // 执行SchemaBuilder
            schemaBuilder.Modify(modifiedTableDescription);
            schemaBuilder.Build();
        }

        // 删除字段，只针对gdb数据库【Delete：删除，Keep：保留】
        public static void DeleteField(string in_table, List<string> fieldNames, string model = "Delete")
        {
            string gdbPath = in_table.TargetWorkSpace();
            string fcName = in_table.TargetFcName();

            using Geodatabase geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));

            TableDefinition tableDefinition = geodatabase.GetDefinition<TableDefinition>(fcName);
            // 创建SchemaBuilder
            SchemaBuilder schemaBuilder = new SchemaBuilder(geodatabase);

            TableDescription tableDescription = new TableDescription(tableDefinition);

            // 从列表中找到要删除的字段
            List<FieldDescription> fieldsToBeRetained = new List<FieldDescription>() { };
            if (model == "Keep")
            {
                IReadOnlyList<Field> fields = tableDefinition.GetFields();
                foreach (Field field in fields)
                {
                    if (field.FieldType == FieldType.Geometry || field.IsEditable == false)     // 不可编辑的字段
                    {
                        continue;
                    }

                    if (fieldNames.Contains(field.Name))
                    {
                        fieldsToBeRetained.Add(new FieldDescription(field));
                    }
                }
            }
            else
            {
                IReadOnlyList<Field> fields = tableDefinition.GetFields();
                foreach (Field field in fields)
                {
                    if (field.FieldType == FieldType.Geometry || field.IsEditable == false)     // 不可编辑的字段
                    {
                        continue;
                    }

                    if (!fieldNames.Contains(field.Name))
                    {
                        fieldsToBeRetained.Add(new FieldDescription(field));
                    }
                }
            }

            // 更新TableDescription
            TableDescription modifiedTableDescription = new TableDescription(tableDescription.Name, fieldsToBeRetained);
            // 执行SchemaBuilder
            schemaBuilder.Modify(modifiedTableDescription);
            schemaBuilder.Build();
        }

        // 更改字段，只针对gdb数据库
        public static void AlterField(string in_table, string fieldName, string newName, int fieldLength = 255, string aliasName = "")
        {
            // 别名默认为字段名，长度为255
            if (aliasName == "")
            {
                aliasName = fieldName;
            }

            string gdbPath = in_table.TargetWorkSpace();
            string fcName = in_table.TargetFcName();

            using Geodatabase geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));

            TableDefinition tableDefinition = geodatabase.GetDefinition<TableDefinition>(fcName);
            // 创建SchemaBuilder
            SchemaBuilder schemaBuilder = new SchemaBuilder(geodatabase);

            TableDescription tableDescription = new TableDescription(tableDefinition);
            // 从列表中找到要删除的字段
            List<FieldDescription> fieldsToBeRetained = new List<FieldDescription>() { };

            IReadOnlyList<Field> fields = tableDefinition.GetFields();
            foreach (Field field in fields)
            {
                if (field.FieldType == FieldType.Geometry || field.IsEditable == false)     // 不可编辑的字段
                {
                    continue;
                }
                if (field.Name != fieldName)
                {
                    fieldsToBeRetained.Add(new FieldDescription(field));
                }
                else
                {
                    //IntPtr intPtr= field.Handle;
                    //FieldInfo fieldInfo;
                    //fieldInfo.name = newName;
                    //fieldInfo.aliasName = aliasName;
                    //fieldInfo.length = fieldLength;
                    //Field field1 = new(intPtr, fieldInfo);

                    //fieldsToBeRetained.Add(fieldDescription);
                }
            }

            // 更新TableDescription
            TableDescription modifiedTableDescription = new TableDescription(tableDescription.Name, fieldsToBeRetained);
            // 执行SchemaBuilder
            schemaBuilder.Modify(modifiedTableDescription);
            schemaBuilder.Build();
        }

        // 更新要素图层的数据源，只针对gdb数据库
        public static void UpdataLayerSource(FeatureLayer featureLayer, string catalogPath)
        {
            CIMDataConnection currentDataConnection = featureLayer.GetDataConnection();

            string connection = System.IO.Path.GetDirectoryName(catalogPath);
            string suffix = System.IO.Path.GetExtension(connection).ToLower();

            var workspaceConnectionString = string.Empty;
            WorkspaceFactory wf = WorkspaceFactory.FileGDB;
            if (suffix == ".sde")
            {
                wf = WorkspaceFactory.SDE;
                var dbGdbConnection = new DatabaseConnectionFile(new Uri(connection, UriKind.Absolute));
                workspaceConnectionString = new Geodatabase(dbGdbConnection).GetConnectionString();
            }
            else
            {
                var dbGdbConnectionFile = new FileGeodatabaseConnectionPath(new Uri(connection, UriKind.Absolute));
                workspaceConnectionString = new Geodatabase(dbGdbConnectionFile).GetConnectionString();
            }

            string dataset = System.IO.Path.GetFileName(catalogPath);

            // 创建CIMStandardDataConnection
            CIMStandardDataConnection updatedDataConnection = new CIMStandardDataConnection()
            {
                WorkspaceConnectionString = workspaceConnectionString,
                WorkspaceFactory = wf,
                Dataset = dataset,
                DatasetType = esriDatasetType.esriDTFeatureClass
            };

            // 更新路径
            featureLayer.SetDataConnection(updatedDataConnection);

            // 释放缓存
            featureLayer.ClearDisplayCache();
        }

        // 从路径或图层中获取字段Field列表。字段类型：
        // allof=>真的全部
        // all=>可编辑的全部
        // notEdit=>不可编辑的全部
        // text=>字符串
        // float=>可编辑的浮点型
        // float_all=>所有的浮点型
        // int=>可编辑的整型
        // int_all=>所有的整型
        // math=>可编辑的数字型
        // math_all=>所有的数字型
        // oid=>OBJECTID
        public static List<Field> GetFieldsFromTarget(object layerName, string field_type = "all")
        {
            List<Field> fields_ori = new List<Field>();     // 所有字段
            List<Field> editFields = new List<Field>();    // 可编辑字段
            List<Field> notEditFields = new List<Field>();   // 不可编辑字段
            // 获取Table
            Table table;
            if (layerName is string)
            {
                table = ((string)layerName).TargetTable();
            }
            else if (layerName is FeatureLayer)
            {
                table = ((FeatureLayer)layerName).GetTable();
            }
            else
            {
                table = null;
            }
            // 获取所有字段
            fields_ori = table.GetDefinition().GetFields().ToList();
            // 可编辑和不可编辑的字段
            foreach (var field in fields_ori)
            {
                if (field.IsEditable && field.FieldType != FieldType.Geometry)
                {
                    editFields.Add(field);
                }
                else
                {
                    notEditFields.Add(field);
                }
            }

            // 输出字段类型
            if (field_type == "allof")  // 真的全部
            {
                return fields_ori;
            }
            else if (field_type == "all")  // 可编辑的全部
            {
                return editFields;
            }
            else if (field_type == "notEdit")    // 不可编辑的全部
            {
                return notEditFields;
            }
            else if (field_type == "text")  // 字符串
            {
                List<Field> text_fields = new List<Field>();
                foreach (Field field in editFields)
                {
                    if (field.FieldType == FieldType.String)
                    {
                        text_fields.Add(field);
                    }
                }
                return text_fields;
            }
            else if (field_type == "float")   // 可编辑的浮点型
            {
                List<Field> float_fields = new List<Field>();
                foreach (Field field in editFields)
                {
                    if (field.FieldType == FieldType.Single || field.FieldType == FieldType.Double)
                    {
                        float_fields.Add(field);
                    }
                }
                return float_fields;
            }

            else if (field_type == "float_all")   // 所有的浮点型
            {
                List<Field> float_fields = new List<Field>();
                foreach (Field field in fields_ori)
                {
                    if (field.FieldType == FieldType.Single || field.FieldType == FieldType.Double)
                    {
                        float_fields.Add(field);
                    }
                }
                return float_fields;
            }

            else if (field_type == "int")   // 可编辑的整型
            {
                List<Field> int_fields = new List<Field>();
                foreach (Field field in editFields)
                {
                    if (field.FieldType == FieldType.SmallInteger || field.FieldType == FieldType.Integer)
                    {
                        int_fields.Add(field);
                    }
                }
                return int_fields;
            }

            else if (field_type == "int_all")   // 所有的整型
            {
                List<Field> int_fields = new List<Field>();
                foreach (Field field in fields_ori)
                {
                    if (field.FieldType == FieldType.SmallInteger || field.FieldType == FieldType.Integer)
                    {
                        int_fields.Add(field);
                    }
                }
                return int_fields;
            }

            else if (field_type == "math")    // 可编辑的数字型
            {
                List<Field> float_fields = new List<Field>();
                foreach (Field field in editFields)
                {
                    if (field.FieldType == FieldType.Single || field.FieldType == FieldType.Double || field.FieldType == FieldType.SmallInteger || field.FieldType == FieldType.Integer)
                    {
                        float_fields.Add(field);
                    }
                }
                return float_fields;
            }

            else if (field_type == "math_all")    // 所有的数字型
            {
                // 获取数字型字段
                List<Field> float_fields = new List<Field>();
                foreach (Field field in fields_ori)
                {
                    if (field.FieldType == FieldType.Single || field.FieldType == FieldType.Double || field.FieldType == FieldType.SmallInteger || field.FieldType == FieldType.Integer)
                    {
                        float_fields.Add(field);
                    }
                }
                return float_fields;
            }

            else if (field_type == "oid")    // OBJECTID
            {
                // 获取数字型字段
                List<Field> oid_fields = new List<Field>();
                foreach (Field field in fields_ori)
                {
                    if (field.FieldType == FieldType.OID)
                    {
                        oid_fields.Add(field);
                    }
                }
                return oid_fields;
            }

            else
            {
                return null;
            }
        }

        // 从路径或图层中获取字段Field列表【字段类型：all=>全部字段，text=>可写的字符串字段，float=>浮点型，int=>整型, notEdit=>不可编辑的】
        public static List<Field> GetFieldsFromTarget(FeatureLayer layer, string field_type = "all")
        {
            List<Field> fields_ori = new List<Field>();
            List<Field> editFields = new List<Field>();
            List<Field> notEditFields = new List<Field>();
            // 获取Table
            Table table = layer.GetTable();
            // 获取所有字段
            fields_ori = table.GetDefinition().GetFields().ToList();
            // 移除不可编辑的字段
            foreach (var field in fields_ori)
            {
                if (field.IsEditable)
                {
                    editFields.Add(field);
                }
                else
                {
                    notEditFields.Add(field);
                }
            }

            // 输出字段类型
            if (field_type == "allof")  // 真的全部
            {
                return fields_ori;
            }
            else if (field_type == "all")  //可编辑的全部
            {
                return editFields;
            }
            else if (field_type == "text")
            {
                // 获取可编辑的字符串字段
                List<Field> text_fields = new List<Field>();
                foreach (Field field in editFields)
                {
                    if (field.FieldType == FieldType.String)
                    {
                        text_fields.Add(field);
                    }
                }
                return text_fields;
            }
            else if (field_type == "float")
            {
                // 获取可编辑的数字型字段
                List<Field> float_fields = new List<Field>();
                foreach (Field field in editFields)
                {
                    if (field.FieldType == FieldType.Single || field.FieldType == FieldType.Double)
                    {
                        float_fields.Add(field);
                    }
                }
                return float_fields;
            }

            else if (field_type == "float_all")
            {
                // 获取可编辑的数字型字段
                List<Field> float_fields = new List<Field>();
                foreach (Field field in fields_ori)
                {
                    if (field.FieldType == FieldType.Single || field.FieldType == FieldType.Double)
                    {
                        float_fields.Add(field);
                    }
                }
                return float_fields;
            }

            else if (field_type == "int")
            {
                // 获取可编辑的数字型字段
                List<Field> int_fields = new List<Field>();
                foreach (Field field in editFields)
                {
                    if (field.FieldType == FieldType.SmallInteger || field.FieldType == FieldType.Integer)
                    {
                        int_fields.Add(field);
                    }
                }
                return int_fields;
            }
            else if (field_type == "notEdit")
            {
                return notEditFields;
            }
            else
            {
                return null;
            }
        }

        // 从路径或图层中按字段名获取字段Field
        public static Field GetFieldFromString(string layerName, string fieldName)
        {
            Field field = null;
            // 获取Table
            Table table = layerName.TargetTable();

            // 获取所有字段
            List<Field> fields_ori = table.GetDefinition().GetFields().ToList();
            // 移除不可编辑的字段
            foreach (Field fd in fields_ori)
            {
                if (fd.Name == fieldName)
                {
                    field = fd;
                }
            }
            // 返回字段
            return field;
        }

        // 从路径或图层中获取字段Field的名称列表【字段类型：all=>全部字段，text=>可写的字符串字段，float=>浮点型，int=>整型】
        public static List<string> GetFieldsNameFromTarget(string layerName, string field_type = "all")
        {
            List<Field> fields = GetFieldsFromTarget(layerName, field_type);
            List<string> names = new List<string>();
            foreach (var field in fields)
            {
                names.Add(field.Name);
            }
            return names;

        }

        // 检查图层中是否包含相应字段， target支持lyName和FeatureLayer
        public static bool IsHaveFieldInTarget(object target, string fieldName)
        {
            bool result = false;

            // 先判断一下，输入是否有空值
            if (target == null || fieldName == null || fieldName == "")
            {
                result = false;
            }

            List<string> str_fields = new List<string>();
            IReadOnlyList<Field> fields = new List<Field>();
            // 获取表
            Table init_table = null;
            if (target is string)
            {
                string lyName = target as string;

                if (lyName != "")
                {
                    init_table = lyName.TargetTable();
                }
            }
            else if (target is FeatureLayer)
            {
                FeatureLayer featureLayer = target as FeatureLayer;
                init_table = featureLayer.GetTable();
            }

            if (init_table == null)
            {
                return false;
            }

            // 获取字段
            fields = init_table.GetDefinition().GetFields();
            // 生成字段列表
            foreach (var item in fields)
            {
                str_fields.Add(item.Name);
            }
            // 查找是否有该字段
            if (str_fields.Contains(fieldName))
            {
                result = true;
            }

            return result;
        }

        // 获取面要素的界线总长度
        public static double GetLineLength(Polygon polygon)
        {
            double len = 0;

            // 获取面要素的部件（内外环）
            var parts = polygon.Parts.ToList();
            foreach (ReadOnlySegmentCollection collection in parts)
            {
                List<MapPoint> points = new List<MapPoint>();
                // 每个环进行处理（第一个为外环，其它为内环）
                foreach (Segment segment in collection)
                {
                    double partLenth = segment.Length;     // 线长度
                    len += partLenth;
                }
            }
            return len;
        }

        // 从路径或图层中获取指定字段的List值
        public static List<string> GetFieldValuesFromPath(string inputData, string inputField)
        {
            List<string> list = new List<string>();
            // 获取Table
            Table table = inputData.TargetTable();
            using (RowCursor rowCursor = table.Search())
            {
                while (rowCursor.MoveNext())
                {
                    using Row row = rowCursor.Current;
                    // 获取value
                    var va = row[inputField];
                    if (va is not null)
                    {
                        // 如果是新值，就加入到list中
                        if (!list.Contains(va.ToString()))
                        {
                            list.Add(va.ToString());
                        }
                    }
                }
            }
            return list;
        }

        // 从路径或图层中获取Dictionary
        public static Dictionary<string, string> GetDictFromPath(string inputData, string in_field_01, string in_field_02)
        {
            Dictionary<string, string> dict = new();
            // 获取Table
            Table table = inputData.TargetTable();
            // 逐行游标
            using (RowCursor rowCursor = table.Search())
            {
                while (rowCursor.MoveNext())
                {
                    using (Row row = rowCursor.Current)
                    {
                        // 获取value
                        var key = row[in_field_01];
                        var value = row[in_field_02];
                        if (key is not null && value is not null)
                        {
                            var va = value.ToString();
                            // 如果没有重复key值，则纳入dict
                            if (!dict.Keys.Contains(key.ToString()))
                            {
                                dict.Add(key.ToString(), va);
                            }
                        }
                    }
                }
            }
            return dict;
        }

        // 从路径或图层中获取Dictionary
        public static Dictionary<string, double> GetDictFromPathDouble(string inputData, string in_field_01, string in_field_02)
        {
            Dictionary<string, double> dict = new();
            // 获取Table
            Table table = inputData.TargetTable();
            // 逐行游标
            using (RowCursor rowCursor = table.Search())
            {
                while (rowCursor.MoveNext())
                {
                    using (Row row = rowCursor.Current)
                    {
                        // 获取value
                        var key = row[in_field_01];
                        var value = row[in_field_02];
                        if (key is not null && value is not null)
                        {
                            var va = double.Parse(value.ToString());
                            // 如果没有重复key值，则纳入dict
                            if (!dict.Keys.Contains(key.ToString()))
                            {
                                dict.Add(key.ToString(), va);
                            }
                        }
                    }
                }
            }
            return dict;
        }


        // 从路径或图层中获取Dictionary【多个字段+double】
        public static Dictionary<List<string>, double> GetDictListFromPathDouble(string inputData, List<string> in_fields, string double_field)
        {
            Dictionary<List<string>, double> dict = new();
            // 获取Table
            Table table = inputData.TargetTable();
            // 逐行游标
            using (RowCursor rowCursor = table.Search())
            {
                while (rowCursor.MoveNext())
                {
                    using Row row = rowCursor.Current;
                    // 获取value
                    List<string> in_values = new List<string>();
                    foreach (var in_field in in_fields)
                    {
                        var key = row[in_field];
                        if (key is not null)
                        {
                            in_values.Add(key.ToString());
                        }
                        else
                        {
                            in_values.Add("");
                        }
                    }
                    var value = row[double_field];
                    if (value is not null)
                    {
                        var va = double.Parse(value.ToString());
                        // 如果没有重复key值，则纳入dict
                        if (!dict.ContainsKey(in_values))
                        {
                            dict.Add(in_values, va);
                        }
                    }
                }
            }
            return dict;
        }

        // 清除GDB里的所有要素，排除文字exText
        public static async void ClearGDBFiles(string gdbPath, string exText = "")
        {
            await QueuedTask.Run(() =>
            {
                using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath))))
                {
                    // 获取文件地理数据库中的所有要素类
                    IReadOnlyList<FeatureClassDefinition> featureClassDefinitions = gdb.GetDefinitions<FeatureClassDefinition>();

                    // 删除每个要素类
                    foreach (var featureClassDefinition in featureClassDefinitions)
                    {
                        string fcName = featureClassDefinition.GetName();

                        using FeatureClass featureClass = gdb.OpenDataset<FeatureClass>(fcName);

                        if (exText != "")
                        {
                            if (!fcName.Contains(exText))
                            {
                                Arcpy.Delect(featureClass);
                            }
                        }
                        else
                        {
                            // 删除要素类中的所有要素
                            Arcpy.Delect(featureClass);
                        }
                    }
                };
            });


        }

        // 通过Geometry创建要素类
        public static void CreateFeatureClassByGeometry(string gdbPath, string fcName, Geometry geometry, SpatialReference sr, string type = "POINT")
        {
            // 使用地理数据库路径创建地理数据库对象
            using Geodatabase geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath)));

            Arcpy.CreateFeatureclass(gdbPath, fcName, type, sr);

            // 打开示例要素
            using FeatureClass enterpriseFeatureClass = geodatabase.OpenDataset<FeatureClass>(fcName);
            // 创建编辑操作对象
            EditOperation editOperation = new EditOperation();

            // 获取要素定义
            FeatureClassDefinition featureClassDefinition = enterpriseFeatureClass.GetDefinition();

            // 创建RowBuffer
            using RowBuffer rowBuffer = enterpriseFeatureClass.CreateRowBuffer();

            // 给新添加的行设置形状
            rowBuffer[featureClassDefinition.GetShapeField()] = geometry;

            // 在表中创建新行
            using Feature feature = enterpriseFeatureClass.CreateRow(rowBuffer);

        }

        // 从打开的属性表中获取Table
        public static Table GetTableFromView()
        {
            // 获取当前激活的表格视图
            var tableView = TableView.Active;
            if (tableView == null || tableView.MapMember == null)
            {
                return null;
            }

            // 获取表格
            Table table;
            if (tableView.MapMember is FeatureLayer featureLayer)
            {
                table = featureLayer.GetTable();
            }
            else if (tableView.MapMember is StandaloneTable standaloneTable)
            {
                table = standaloneTable.GetTable();
            }
            else
            {
                table = null;
            }

            return table;
        }

        // 从打开的属性表中获取Table
        public static Field GetSelectField()
        {
            // 获取选定的字段名
            var tableView = TableView.Active;
            string selectedField = tableView.GetSelectedFields().FirstOrDefault();
            // 获取Table
            Table table = GetTableFromView();
            // 获取Field
            Field field = table.GetDefinition().GetFields().FirstOrDefault(f => f.Name == selectedField);

            return field;
        }

    }
}
