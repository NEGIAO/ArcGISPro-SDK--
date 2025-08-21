using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FieldDescription = ArcGIS.Core.Data.DDL.FieldDescription;

namespace CCTool.Scripts.ToolManagers
{
    // 综合型工具，复杂方法的提取复用
    public class ComboTool
    {
        // 汇总统计，按【单字段】分类汇总面积，包括合计值
        public static Dictionary<string, double> StatisticsPlus(string in_table, string staField, string areaField, string total_field = "", double xs = 1, string sql = null)
        {
            Dictionary<string, double> dict = new();
            Table table = in_table.TargetTable();

            var queryFilter = new QueryFilter { WhereClause = sql };

            using (RowCursor rowCursor = table.Search(queryFilter, false))
            {
                while (rowCursor.MoveNext())
                {
                    using Row row = rowCursor.Current;
                    var sta = row[staField];
                    var area = row[areaField];
                    if (sta != null && area != null)
                    {
                        string bj = sta.ToString();
                        double mj = (double)area / xs;
                        // 加入dict
                        if (!dict.ContainsKey(bj))
                        {
                            dict[bj] = mj;
                        }
                        else
                        {
                            dict[bj] += mj;
                        }
                        // 计算合计值
                        if (total_field != "")
                        {
                            if (!dict.ContainsKey(total_field))
                            {
                                dict[total_field] = mj;
                            }
                            else
                            {
                                dict[total_field] += mj;
                            }
                        }
                    }
                }
            }

            return dict;
        }

        // 汇总统计，按【多字段】分类汇总面积，包括合计值
        public static Dictionary<string, double> StatisticsPlus(string in_table, List<string> staFields, string areaField, string total_field = "", double xs = 1, string sql = null)
        {
            Dictionary<string, double> dict = new();
            Table table = in_table.TargetTable();

            var queryFilter = new QueryFilter { WhereClause = sql };

            using (RowCursor rowCursor = table.Search(queryFilter, false))
            {
                while (rowCursor.MoveNext())
                {
                    using Row row = rowCursor.Current;
                    var area = row[areaField];
                    if (area != null)
                    {
                        double mj = double.Parse(area.ToString()) / xs;
                        foreach (string staField in staFields)
                        {
                            var sta = row[staField];
                            if (sta != null)
                            {
                                string bj = sta.ToString();

                                // 加入dict
                                if (!dict.ContainsKey(bj))
                                {
                                    dict[bj] = mj;
                                }
                                else
                                {
                                    dict[bj] += mj;
                                }

                            }
                        }

                        // 计算合计值
                        if (total_field != "")
                        {
                            if (!dict.ContainsKey(total_field))
                            {
                                dict[total_field] = mj;
                            }
                            else
                            {
                                dict[total_field] += mj;
                            }
                        }

                    }
                }
            }

            return dict;
        }

        // 创建四个方位点
        public static void Create4Point(string in_data, string out_data, string bjField)
        {
            string out_gdb = out_data.TargetWorkSpace();
            string bjd = out_data.TargetFcName();

            // 判断一下是否存在目标要素，如果有的话，就删掉重建
            bool isHaveTarget = out_gdb.IsHaveFeaturClass(bjd);
            if (isHaveTarget)
            {
                Arcpy.Delect($@"{out_gdb}\{bjd}");
            }

            // 获取目标FeatureLayer
            FeatureClass featureClass = in_data.TargetFeatureClass();
            // 获取坐标系
            SpatialReference sr = featureClass.GetDefinition().GetSpatialReference();

            List<PointAtt> pts = new List<PointAtt>();

            // 遍历面要素类中的所有要素
            RowCursor cursor = featureClass.Search();
            while (cursor.MoveNext())
            {
                using var feature = cursor.Current as Feature;
                string dkName = bjField == "" ? "" : feature[bjField]?.ToString();
                // 获取要素的几何
                Polygon polygon = feature.GetShape() as Polygon;

                double xx = (polygon.Extent.XMax + polygon.Extent.XMin) / 2;
                double yy = (polygon.Extent.YMax + polygon.Extent.YMin) / 2;
                // 左
                PointAtt pt_left = new PointAtt
                {
                    Name = dkName,
                    Des = "左",
                    X = polygon.Extent.XMin,
                    Y = yy
                };
                // 右
                PointAtt pt_right = new PointAtt
                {
                    Name = dkName,
                    Des = "右",
                    X = polygon.Extent.XMax,
                    Y = yy
                };
                // 上
                PointAtt pt_up = new PointAtt
                {
                    Name = dkName,
                    Des = "上",
                    X = xx,
                    Y = polygon.Extent.YMax,
                };
                // 下
                PointAtt pt_bottom = new PointAtt
                {
                    Name = dkName,
                    Des = "下",
                    X = xx,
                    Y = polygon.Extent.YMin,
                };

                pts.Add(pt_left);
                pts.Add(pt_right);
                pts.Add(pt_up);
                pts.Add(pt_bottom);
            }


            // 创建一个ShapeDescription
            var shapeDescription = new ShapeDescription(GeometryType.Point, sr)
            {
                HasM = false,
                HasZ = false
            };
            // 定义字段
            var name = new ArcGIS.Core.Data.DDL.FieldDescription("Name", FieldType.String);
            var des = new ArcGIS.Core.Data.DDL.FieldDescription("方位", FieldType.String);

            // 打开数据库gdb
            using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(out_gdb))))
            {
                // 收集字段列表
                var fieldDescriptions = new List<ArcGIS.Core.Data.DDL.FieldDescription>()
                        {
                            name,des
                        };

                // 创建FeatureClassDescription
                var fcDescription = new FeatureClassDescription(bjd, fieldDescriptions, shapeDescription);
                // 创建SchemaBuilder
                SchemaBuilder schemaBuilder = new SchemaBuilder(gdb);
                // 将创建任务添加到DDL任务列表中
                schemaBuilder.Create(fcDescription);
                // 执行DDL
                bool success = schemaBuilder.Build();

                // 创建要素并添加到要素类中
                using FeatureClass fc = gdb.OpenDataset<FeatureClass>(bjd);
                /// 构建点要素
                // 创建编辑操作对象
                EditOperation editOperation = new EditOperation();
                editOperation.Callback(context =>
                {
                    // 获取要素定义
                    FeatureClassDefinition featureClassDefinition = fc.GetDefinition();
                    // 循环创建点
                    foreach (var pt in pts)
                    {
                        // 创建RowBuffer
                        using RowBuffer rowBuffer = fc.CreateRowBuffer();

                        // 写入字段值
                        rowBuffer["Name"] = pt.Name;
                        rowBuffer["方位"] = pt.Des;

                        // 创建点几何
                        MapPointBuilderEx mapPointBuilderEx = new(pt.X, pt.Y);
                        // 给新添加的行设置形状
                        rowBuffer[featureClassDefinition.GetShapeField()] = mapPointBuilderEx.ToGeometry();

                        // 在表中创建新行
                        using Feature feature = fc.CreateRow(rowBuffer);
                        context.Invalidate(feature);      // 标记行为无效状态
                    }

                }, fc);

                // 执行编辑操作
                editOperation.Execute();
            }

            // 保存
            Project.Current.SaveEditsAsync();


        }

        // 更新字段值（数字型，增加值）
        public static string IncreRowValueToTable(string in_table, string insert_value)
        {
            // 获取字段和值的列表
            List<string> keyAndValues = insert_value.Split(";").ToList();
            // 获取Table
            using Table sta_table = in_table.TargetTable();
            // 判断是不是已有该字段
            bool isHaveRow = false;
            string firstKey = keyAndValues[0].Split(",")[0];
            string firstValue = keyAndValues[0].Split(",")[1];
            using (RowCursor rowCursor = sta_table.Search())
            {
                while (rowCursor.MoveNext())
                {
                    using Row row = rowCursor.Current;

                    // 获取value
                    var va = row[firstKey];
                    if (va is not null)
                    {
                        if (va.ToString() == firstValue)
                        {
                            isHaveRow = true;
                        }
                    }
                }
            }
            // 获取表定义
            TableDefinition tableDefinition = sta_table.GetDefinition();

            // 如果有符合该行的字段，则更新字段值
            if (isHaveRow)
            {
                using RowCursor rowCursor2 = sta_table.Search(null, false);
                while (rowCursor2.MoveNext())
                {
                    using Row row2 = rowCursor2.Current;
                    // 获取value
                    var va = row2[firstKey];
                    // 如果符合
                    if (va is not null)
                    {
                        if (va.ToString() == firstValue)
                        {
                            // 写入字段值
                            foreach (var keyAndValue in keyAndValues)
                            {
                                string key = keyAndValue.Split(",")[0];
                                string value = keyAndValue.Split(",")[1];
                                if (key != firstKey)
                                {
                                    double double_value = double.Parse(row2[key].ToString());
                                    row2[key] = double_value + double.Parse(value);
                                }
                            }
                        }
                    }
                    row2.Store();
                }
            }
            // 如果没有符合该行的字段，则插入并更新
            else
            {
                EditOperation editOperation = new EditOperation();  // 创建编辑操作对象
                // 创建RowBuffer
                using RowBuffer rowBuffer = sta_table.CreateRowBuffer();
                // 写入字段值
                foreach (var keyAndValue in keyAndValues)
                {
                    string key = keyAndValue.Split(",")[0];
                    string value = keyAndValue.Split(",")[1];
                    rowBuffer[key] = value;
                }
                // 在表中创建新行
                using Row row = sta_table.CreateRow(rowBuffer);
                editOperation.Execute();
            }
            // 返回值
            return in_table;
        }

        // 在表中插入行并赋值字符串
        public static string UpdataRowToTable(string in_table, string insert_value)
        {
            // 获取字段和值的列表
            List<string> keyAndValues = insert_value.Split(";").ToList();
            // 获取Table
            using Table sta_table = in_table.TargetTable();
            // 判断是不是已有该字段
            bool isHaveRow = false;
            string firstKey = keyAndValues[0].Split(",")[0];
            string firstValue = keyAndValues[0].Split(",")[1];
            using (RowCursor rowCursor = sta_table.Search())
            {
                while (rowCursor.MoveNext())
                {
                    using Row row = rowCursor.Current;

                    // 获取value
                    var va = row[firstKey];
                    if (va is not null)
                    {
                        if (va.ToString() == firstValue)
                        {
                            isHaveRow = true;
                        }
                    }
                }
            }
            // 获取表定义
            TableDefinition tableDefinition = sta_table.GetDefinition();

            // 如果有符合该行的字段，则更新字段值
            if (isHaveRow)
            {
                using RowCursor rowCursor2 = sta_table.Search();
                while (rowCursor2.MoveNext())
                {
                    using Row row2 = rowCursor2.Current;
                    // 获取value
                    var va = row2[firstKey];
                    // 如果符合
                    if (va.ToString() == firstValue)
                    {
                        // 写入字段值
                        foreach (var keyAndValue in keyAndValues)
                        {
                            string key = keyAndValue.Split(",")[0];
                            string value = keyAndValue.Split(",")[1];
                            row2[key] = value;
                        }
                    }
                    row2.Store();
                }
            }
            // 如果没有符合该行的字段，则插入并更新
            else
            {
                EditOperation editOperation = new EditOperation();  // 创建编辑操作对象
                // 创建RowBuffer
                using RowBuffer rowBuffer = sta_table.CreateRowBuffer();
                // 写入字段值
                foreach (var keyAndValue in keyAndValues)
                {
                    string key = keyAndValue.Split(",")[0];
                    string value = keyAndValue.Split(",")[1];
                    rowBuffer[key] = value;
                }
                // 在表中创建新行
                using Row row = sta_table.CreateRow(rowBuffer);
                editOperation.Execute();
            }
            // 返回值
            return in_table;
        }

        // 拓扑检查
        public static string TopologyCheck(string in_data_path, List<string> rules, string outGDBPath)
        {
            string in_data = in_data_path.Replace(@"/", @"\");  // 兼容两种符号
            string db_name = "Top2Check";    // 要素数据集名
            string fc_name = "top_fc";        // 要素名
            string top_name = "Topology";       // TOP名
            string db_path = outGDBPath + "\\" + db_name;    // 要素数据集路径
            string fc_top_path = db_path + "\\" + fc_name;         // 要素路径
            string top_path = db_path + "\\" + top_name;         // TOP路径

            string in_gdb = in_data[..(in_data.IndexOf(@".gdb") + 4)];
            string in_fc = in_data[(in_data.LastIndexOf(@"\") + 1)..];

            // 打开GDB数据库
            using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(in_gdb)));
            // 获取要素类
            FeatureClassDefinition featureClasse = gdb.GetDefinition<FeatureClassDefinition>(in_fc);
            //获取图层的坐标系
            var sr = featureClasse.GetSpatialReference();
            //在数据库中创建要素数据集
            Arcpy.CreateFeatureDataset(outGDBPath, db_name, sr);
            // 将所选要素复制到创建的要素数据集中
            Arcpy.CopyFeatures(in_data, fc_top_path);
            // 新建拓扑
            Arcpy.CreateTopology(db_path, top_name);
            // 向拓扑中添加要素
            Arcpy.AddFeatureClassToTopology(top_path, fc_top_path);
            // 添加拓扑规则
            foreach (var rule in rules)
            {
                Arcpy.AddRuleToTopology(top_path, rule, fc_top_path);
            }
            // 验证拓扑
            Arcpy.ValidateTopology(top_path);
            // 输出TOP错误
            Arcpy.ExportTopologyErrors(top_path, outGDBPath, "TopErr");
            // 删除中间数据
            Arcpy.Delect(db_path);
            // 返回值
            return outGDBPath;
        }

        // 要素消除
        public static string FeatureClassEliminate(string in_fc, string out_fc, string sql, string ex_sql = "")
        {
            string layer = "待消除要素";
            string el_layer = layer;
            // 创建要素图层
            Arcpy.MakeFeatureLayer(in_fc, layer, true);
            // 按属性选择图层
            Arcpy.SelectLayerByAttribute(layer, sql);
            // 消除
            Arcpy.Eliminate(el_layer, out_fc, ex_sql);
            // 移除图层
            MapCtlTool.RemoveLayer(layer);
            // 返回值
            return out_fc;
        }

        // 获取面空洞【输出模式：空洞 | 外边界】
        public static string GetCave(string in_featureClass, string out_featureClass, string model = "空洞")
        {
            // 获取默认数据库
            var gdb = Project.Current.DefaultGeodatabasePath;
            // 融合要素
            Arcpy.Dissolve(in_featureClass, gdb + @"\dissolve_fc");
            // 面转线
            Arcpy.PolygonToLine(gdb + @"\dissolve_fc", gdb + @"\dissolve_line");
            // 要素转面
            Arcpy.FeatureToPolygon(gdb + @"\dissolve_line", gdb + @"\dissolve_polygon");
            // 再融合，获取边界
            Arcpy.Dissolve(gdb + @"\dissolve_polygon", gdb + @"\dissolve_fin");
            // 擦除，获取空洞
            Arcpy.Erase(gdb + @"\dissolve_fin", gdb + @"\dissolve_fc", gdb + @"\single_fc");
            // 单部件转多部件，输出
            if (model == @"空洞")
            {
                Arcpy.MultipartToSinglepart(gdb + @"\single_fc", out_featureClass);
            }
            else if (model == @"外边界")
            {
                Arcpy.MultipartToSinglepart(gdb + @"\dissolve_fin", out_featureClass);
            }
            // 删除中间要素
            List<string> list_fc = new List<string>() { "dissolve_fc", "dissolve_line", "dissolve_polygon", "dissolve_fin", "single_fc" };
            foreach (var fc in list_fc)
            {
                Arcpy.Delect(gdb + @"\" + fc);
            }
            // 返回值
            return out_featureClass;
        }

        // 按面积平差计算
        public static string AdjustmentByArea(string yd, string field, string area_type, int digit, double totalArea, string clipfc_sort)
        {
            string def_gdb = Project.Current.DefaultGeodatabasePath;
            string area_line = def_gdb + @"\area_line";

            string clipfc_sta = def_gdb + @"\clipfc_sta";
            string clipfc_updata = def_gdb + @"\clipfc_updata";

            if (area_type == "投影")
            {
                Arcpy.CalculateField(yd, field, $"round(!shape_area!,{digit})");
                Arcpy.Statistics(yd, clipfc_sta, $"{field} SUM", "");          // 汇总
            }
            else if (area_type == "图斑")
            {
                Arcpy.CalculateField(yd, field, $"round(!shape.geodesicarea!,{digit})");
                Arcpy.Statistics(yd, clipfc_sta, $"{field} SUM", "");          // 汇总
            }

            // 获取投影面积，图斑面积
            double mj_fc = double.Parse(clipfc_sta.TargetCellValue($"SUM_{field}", ""));

            // 面积差值
            double dif_mj = Math.Round(Math.Round(totalArea, digit) - Math.Round(mj_fc, digit), digit);

            // 空间连接，找出变化图斑（即需要平差的图斑）
            string area = $@"{def_gdb}\area";
            Arcpy.Dissolve(yd, area);
            Arcpy.FeatureToLine(area, area_line);
            if (GisTool.IsHaveFieldInTarget(area_line, "BJM"))
            {
                Arcpy.DeleteField(area_line, "BJM", "KEEP_FIELDS");
            }
            Arcpy.SpatialJoin(yd, area_line, clipfc_updata);
            Arcpy.AddField(clipfc_updata, "平差", "TEXT");
            Arcpy.CalculateField(clipfc_updata, "平差", "''");
            // 排序
            Arcpy.Sort(clipfc_updata, clipfc_sort, "Shape_Area DESCENDING", "UR");
            double area_total = 0;

            // 获取Table
            using Table table = clipfc_sort.TargetTable();

            // 汇总变化图斑的面积
            using (RowCursor rowCursor = table.Search())
            {
                while (rowCursor.MoveNext())
                {
                    using Row row = rowCursor.Current;
                    var va = int.Parse(row["Join_Count"].ToString());
                    if (va == 1)     // 如果是变化图斑
                    {
                        area_total += double.Parse(row[field].ToString());
                    }
                }
            }
            // 第一轮平差
            double area_pc_1 = 0;
            using (RowCursor rowCursor1 = table.Search())
            {
                while (rowCursor1.MoveNext())
                {
                    using Row row = rowCursor1.Current;
                    var va = int.Parse(row["Join_Count"].ToString());
                    if (va == 1)
                    {
                        double area_1 = double.Parse(row[field].ToString());
                        // 单个图斑需要平差的值
                        double area_pc = Math.Round(area_1 / area_total * dif_mj, digit);
                        area_pc_1 += area_pc;
                        // 面积平差
                        row[field] = area_1 + area_pc;

                        row.Store();
                    }
                }
            }
            // 计算剩余平差面积，进行第二轮平差
            double area_total_next = Math.Round(dif_mj - area_pc_1, digit);
            using (RowCursor rowCursor2 = table.Search())
            {
                while (rowCursor2.MoveNext())
                {
                    using Row row = rowCursor2.Current;
                    // 最小平差值
                    double diMin = Math.Round(Math.Pow(0.1, digit), digit);

                    var va = int.Parse(row["Join_Count"].ToString());
                    if (va == 1)
                    {
                        double area_2 = double.Parse(row[field].ToString());
                        // 面积平差
                        if (area_total_next > 0)
                        {
                            row[field] = area_2 + diMin;
                            area_total_next -= diMin;
                        }
                        else if (area_total_next < 0)
                        {
                            row[field] = area_2 - diMin;
                            area_total_next += diMin;
                        }
                        row.Store();
                    }
                }
            }
            // 删除中间要素
            List<string> all = new List<string>() { "area_line", "clipfc", "area", "clipfc_sta", "clipfc_updata" };
            foreach (var item in all)
            {
                Arcpy.Delect(def_gdb + @"\" + item);
            }
            // 返回值
            return clipfc_sort;
        }

        // 裁剪平差计算
        public static string Adjustment(string yd, string area, string clipfc_sort, string area_type = "投影", string unit = "平方米", int digit = 2, string areaField = "")
        {
            string def_gdb = Project.Current.DefaultGeodatabasePath;
            string area_line = def_gdb + @"\area_line";
            string clipfc = def_gdb + @"\clipfc";
            string clipfc_sta = def_gdb + @"\clipfc_sta";
            string clipfc_updata = def_gdb + @"\clipfc_updata";

            // 单位系数设置
            double unit_xs = 0;
            if (unit == "平方米") { unit_xs = 1; }
            else if (unit == "公顷") { unit_xs = 10000; }
            else if (unit == "平方公里") { unit_xs = 1000000; }
            else if (unit == "亩") { unit_xs = 666.66667; }

            // 计算图斑的投影面积和图斑面积
            Arcpy.Clip(yd, area, clipfc);

            if (areaField != "")
            {
                area_type = "图斑";

                Arcpy.AddField(clipfc, area_type, "DOUBLE");
                Arcpy.AddField(area, area_type, "DOUBLE");

                Arcpy.CalculateField(clipfc, area_type, $"!{areaField}!");
                Arcpy.Statistics(clipfc, clipfc_sta, area_type, "");          // 汇总
                // 计算范围的投影面积和图斑面积
                Arcpy.CalculateField(area, area_type, $"round(!shape.geodesicarea!/{unit_xs},{digit})");

            }
            else
            {
                Arcpy.AddField(clipfc, area_type, "DOUBLE");
                Arcpy.AddField(area, area_type, "DOUBLE");

                if (area_type == "投影")
                {
                    Arcpy.CalculateField(clipfc, area_type, $"round(!shape_area!/{unit_xs},{digit})");
                    Arcpy.Statistics(clipfc, clipfc_sta, area_type, "");          // 汇总
                    // 计算范围的投影面积和图斑面积
                    Arcpy.CalculateField(area, area_type, $"round(!shape_area!/{unit_xs},{digit})");
                }
                else if (area_type == "图斑")
                {
                    Arcpy.CalculateField(clipfc, area_type, $"round(!shape.geodesicarea!/{unit_xs},{digit})");
                    Arcpy.Statistics(clipfc, clipfc_sta, area_type, "");          // 汇总
                    // 计算范围的投影面积和图斑面积
                    Arcpy.CalculateField(area, area_type, $"round(!shape.geodesicarea!/{unit_xs},{digit})");
                }
            }

            // 获取投影面积，图斑面积
            double mj_fc = double.Parse(clipfc_sta.TargetCellValue($"SUM_{area_type}", ""));
            double mj_area = double.Parse(area.TargetCellValue(area_type, ""));

            // 面积差值
            double dif_mj = Math.Round(Math.Round(mj_area, digit) - Math.Round(mj_fc, digit), digit);

            // 空间连接，找出变化图斑（即需要平差的图斑）
            Arcpy.FeatureToLine(area, area_line);
            if (GisTool.IsHaveFieldInTarget(area_line, "BJM"))
            {
                Arcpy.DeleteField(area_line, "BJM", "KEEP_FIELDS");
            }
            Arcpy.SpatialJoin(clipfc, area_line, clipfc_updata);
            Arcpy.AddField(clipfc_updata, "平差", "TEXT");
            Arcpy.CalculateField(clipfc_updata, "平差", "''");
            // 排序
            Arcpy.Sort(clipfc_updata, clipfc_sort, "Shape_Area DESCENDING", "UR");
            double area_total = 0;

            // 获取Table
            using Table table = clipfc_sort.TargetTable();

            // 汇总变化图斑的面积
            using (RowCursor rowCursor = table.Search())
            {
                while (rowCursor.MoveNext())
                {
                    using Row row = rowCursor.Current;
                    var va = int.Parse(row["Join_Count"].ToString());
                    if (va == 1)     // 如果是变化图斑
                    {
                        area_total += double.Parse(row[area_type].ToString());
                    }
                }
            }
            // 第一轮平差
            double area_pc_1 = 0;
            using (RowCursor rowCursor1 = table.Search())
            {
                while (rowCursor1.MoveNext())
                {
                    using Row row = rowCursor1.Current;
                    var va = int.Parse(row["Join_Count"].ToString());
                    if (va == 1)
                    {
                        double area_1 = double.Parse(row[area_type].ToString());
                        // 单个图斑需要平差的值
                        double area_pc = Math.Round(area_1 / area_total * dif_mj, digit);
                        area_pc_1 += area_pc;
                        // 面积平差
                        row[area_type] = area_1 + area_pc;

                        row.Store();
                    }
                }
            }
            // 计算剩余平差面积，进行第二轮平差
            double area_total_next = Math.Round(dif_mj - area_pc_1, digit);
            using (RowCursor rowCursor2 = table.Search())
            {
                while (rowCursor2.MoveNext())
                {
                    using Row row = rowCursor2.Current;
                    // 最小平差值
                    double diMin = Math.Round(Math.Pow(0.1, digit), digit);

                    var va = int.Parse(row["Join_Count"].ToString());
                    if (va == 1)
                    {
                        double area_2 = double.Parse(row[area_type].ToString());
                        // 面积平差
                        if (area_total_next > 0)
                        {
                            row[area_type] = area_2 + diMin;
                            area_total_next -= diMin;
                        }
                        else if (area_total_next < 0)
                        {
                            row[area_type] = area_2 - diMin;
                            area_total_next += diMin;
                        }
                        row.Store();
                    }
                }
            }
            // 删除中间要素
            List<string> all = new List<string>() { "area_line", "clipfc", "clipfc_sta", "clipfc_updata" };
            foreach (var item in all)
            {
                Arcpy.Delect(def_gdb + @"\" + item);
            }
            // 返回值
            return clipfc_sort;
        }


        // 裁剪并计算面积
        public static string AdjustmentNot(string yd, string area, string clipfc_sort, string area_type = "投影", string unit = "平方米", int digit = 2)
        {
            // 单位系数设置
            double unit_xs = 0;
            if (unit == "平方米") { unit_xs = 1; }
            else if (unit == "公顷") { unit_xs = 10000; }
            else if (unit == "平方公里") { unit_xs = 1000000; }
            else if (unit == "亩") { unit_xs = 666.66667; }

            // 计算图斑的投影面积和图斑面积
            Arcpy.Clip(yd, area, clipfc_sort);

            Arcpy.AddField(clipfc_sort, area_type, "DOUBLE");

            if (area_type == "投影")
            {
                Arcpy.CalculateField(clipfc_sort, area_type, $"round(!shape_area!/{unit_xs},{digit})");
            }
            else if (area_type == "图斑")
            {
                Arcpy.CalculateField(clipfc_sort, area_type, $"round(!shape.geodesicarea!/{unit_xs},{digit})");
            }
            // 返回值
            return yd;
        }

        // 分区平差计算
        public static string AreaAdjustment(string zoomDissolve, string targetFc, string nameField, string areaType = "投影", string unit = "平方米", int digit = 2)
        {
            // 单位系数设置
            double unit_xs = unit switch
            {
                "平方米" => 1,
                "公顷" => 10000,
                "平方公里" => 1000000,
                "亩" => 666.66667,
                _ => 1,
            };

            // 添加面积字段并计算
            string mjField = "SDMJ";
            GisTool.AddField(targetFc, mjField, FieldType.Double);

            // 复制一个分区要素
            GisTool.AddField(zoomDissolve, mjField, FieldType.Double);

            if (areaType == "投影")
            {
                Arcpy.CalculateField(targetFc, mjField, $"round(!shape_area!/{unit_xs},{digit})");
                Arcpy.CalculateField(zoomDissolve, mjField, $"round(!shape_area!/{unit_xs},{digit})");
            }
            else
            {
                Arcpy.CalculateField(targetFc, mjField, $"round(!shape.geodesicarea!/{unit_xs},{digit})");
                Arcpy.CalculateField(zoomDissolve, mjField, $"round(!shape.geodesicarea!/{unit_xs},{digit})");
            }

            // 按面积排序，方便后面平差
            var defGDB = Project.Current.DefaultGeodatabasePath;
            string SortFc = @$"{defGDB}\SortFc";
            Arcpy.Sort(targetFc, SortFc, $"{mjField} DESCENDING", "UR");

            // 标记名
            List<string> zoomNames = GisTool.GetFieldValuesFromPath(SortFc, nameField);

            foreach (string zoomName in zoomNames)
            {
                // 计算总面积
                string tatolField = "合计";
                string sql = $"{nameField} = '{zoomName}'";
                var fcDict = ComboTool.StatisticsPlus(SortFc, nameField, mjField, tatolField, 1, sql);
                var zoomDict = ComboTool.StatisticsPlus(zoomDissolve, nameField, mjField, tatolField, 1, sql);

                double fcTotal = fcDict[tatolField];
                double zoomTotal = zoomDict[tatolField];

                // 面积差值
                double difMJ = Math.Round(Math.Round(zoomTotal, digit) - Math.Round(fcTotal, digit), digit);

                // 第一轮平差
                QueryFilter queryFilter = new QueryFilter { WhereClause = sql };
                using Table table = SortFc.TargetTable();
                double area_pc_1 = 0;
                using (RowCursor rowCursor1 = table.Search(queryFilter, false))
                {
                    while (rowCursor1.MoveNext())
                    {
                        using Row row = rowCursor1.Current;
                        double area_1 = (double)row[mjField];
                        // 单个图斑需要平差的值
                        double area_pc = Math.Round(area_1 / zoomTotal * difMJ, digit);
                        // 无需调整的图斑就跳过
                        if (area_pc == 0) { continue; }

                        // 要调整就进行面积平差
                        area_pc_1 += area_pc;
                        row[mjField] = area_1 + area_pc;

                        row.Store();
                    }
                }
                // 计算剩余平差面积，进行第二轮平差
                double area_total_next = Math.Round(difMJ - area_pc_1, digit);
                using RowCursor rowCursor2 = table.Search(queryFilter, false);
                while (rowCursor2.MoveNext())
                {
                    using Row row = rowCursor2.Current;
                    // 最小平差值
                    double diMin = Math.Round(Math.Pow(0.1, digit), digit);

                    double area_2 = double.Parse(row[mjField].ToString());
                    // 面积平差
                    if (area_total_next > 0)
                    {
                        row[mjField] = area_2 + diMin;
                        area_total_next -= diMin;
                        row.Store();
                    }
                    else if (area_total_next < 0)
                    {
                        row[mjField] = area_2 - diMin;
                        area_total_next += diMin;
                        row.Store();
                    }
                    else
                    {
                        break;
                    }
                }

            }

            return SortFc;
        }

        // 分区但不进行平差计算，只计算三调面积
        public static string AreaAdjustmentNot( string targetFc,  string areaType = "投影", string unit = "平方米", int digit = 2)
        {
            // 单位系数设置
            double unit_xs = unit switch
            {
                "平方米" => 1,
                "公顷" => 10000,
                "平方公里" => 1000000,
                "亩" => 666.66667,
                _ => 1,
            };

            // 添加面积字段并计算
            string mjField = "SDMJ";
            GisTool.AddField(targetFc, mjField, FieldType.Double);

            if (areaType == "投影")
            {
                Arcpy.CalculateField(targetFc, mjField, $"round(!shape_area!/{unit_xs},{digit})");
            }
            else
            {
                Arcpy.CalculateField(targetFc, mjField, $"round(!shape.geodesicarea!/{unit_xs},{digit})");
            }

            return targetFc;
        }


        // 分解汇总表
        public static Dictionary<string, double> DecomposeSummary(Dictionary<string, double> in_dic)
        {
            // 复制一个dic
            Dictionary<string, double> dic = new Dictionary<string, double>();

            foreach (var item in in_dic)
            {
                // 分不同情况处理
                if (item.Key.Contains("+"))     // 如果是混合用地
                {
                    dic.Add(item.Key, item.Value);
                    // 分析一下是不是纯混合用地
                    string key1 = item.Key[..item.Key.IndexOf("+")];
                    string key2 = item.Key[(item.Key.IndexOf("+") + 1)..];
                    if (key1[..2] != "09" || key2[..2] != "09")
                    {
                        dic.Accumulation("HH", item.Value);
                    }
                    else    //  如果都是商业混合，那就算是商业用地
                    {
                        dic.Accumulation("09", item.Value);
                    }
                }
                else    // 如果不是混合用地
                {
                    // 分解BM
                    List<string> keyList = item.Key.DecomposeBM();
                    foreach (var k in keyList)
                    {
                        dic.Accumulation(k, item.Value);
                    }
                }
            }
            return dic;
        }


        // 局部汇总
        public static Dictionary<string, double> PartSummary(Dictionary<string, double> in_dic, Dictionary<string, List<string>> pairs)
        {
            Dictionary<string, double> result = new Dictionary<string, double>(in_dic);

            foreach (var pair in pairs)
            {
                // 要汇总的列表
                string bj = pair.Key;
                List<string> keyList = pair.Value;

                foreach (var item in in_dic)
                {
                    // 更新建设用地
                    if (keyList.Contains(item.Key))
                    {
                        result.Accumulation(bj, item.Value);
                    }
                }
            }

            // 返回
            return result;
        }


        // 汇总统计加强版Dic
        public static Dictionary<string, double> MultiStatisticsToDic(string in_table, string statistics_field, List<string> case_fields, string total_field = "", double unit = 1)
        {
            Dictionary<string, double> dic = new Dictionary<string, double>();
            // 中间用于计算的汇总表
            string def_gdb = Project.Current.DefaultGeodatabasePath;
            string out_table = @$"{def_gdb}\stb";
            string out_table2 = @$"{def_gdb}\stb2";
            // 循环列表统计
            for (int i = 0; i < case_fields.Count; i++)
            {
                // 调用GP工具【汇总】
                Arcpy.Statistics(in_table, out_table, $"{statistics_field} SUM", case_fields[i]);
                // 生成dic
                Dictionary<string, double> re = GisTool.GetDictFromPathDouble(out_table, case_fields[i], $"SUM_{statistics_field}");

                // 合并dic
                dic = dic.Union(re).ToDictionary(kv => kv.Key, kv => kv.Value);
            }

            Dictionary<string, double> dict = new Dictionary<string, double>(dic);
            // 如果要统计总值
            if (total_field != "")
            {
                Arcpy.Statistics(in_table, out_table2, $"{statistics_field} SUM", "");
                List<string> result = out_table2.GetFieldValues($"SUM_{statistics_field}");
                double sumNumber = result.Count > 0 ? double.Parse(result[0]) : 0;
                dict.Add(total_field, sumNumber);
            }
            // 单位修改
            if (unit != 1)
            {
                foreach (var item in dict)
                {
                    double va = item.Value;
                    dict[item.Key] = va / unit;
                }
            }

            // 删除中间表
            Arcpy.Delect(out_table);
            Arcpy.Delect(out_table2);
            // 返回
            return dict;
        }


        // 汇总统计加强版
        public static void MultiStatistics(string in_table, string out_table, string statistics_fields, List<string> case_fields, string total_field = "合计", int unit = 0, bool is_output = false)
        {
            List<string> list_table = new List<string>();
            for (int i = 0; i < case_fields.Count; i++)
            {
                Arcpy.Statistics(in_table, out_table + i.ToString(), statistics_fields, case_fields[i]);    // 调用GP工具【汇总】
                Arcpy.AlterField(out_table + i.ToString(), case_fields[i], @"分组", @"分组");  // 调用GP工具【更改字段】
                list_table.Add(out_table + i.ToString());
            }
            Arcpy.Statistics(in_table, out_table + "_total", statistics_fields, "");    // // 调用GP工具【汇总】
            Arcpy.AddField(out_table + "_total", @"分组", "TEXT");    // 调用GP工具【更改字段】
            Arcpy.CalculateField(out_table + "_total", @"分组", "\"" + total_field + "\"");    // 调用GP工具【计算字段】
            list_table.Add(out_table + "_total");     // 加入列表
            // 合并汇总表
            Arcpy.Merge(list_table, out_table, is_output);       // 调用GP工具【合并】
                                                                 // 单位转换
            if (unit > 0)
            {
                string fd = "SUM_" + statistics_fields.Replace(" SUM", "");
                FieldCalTool.ChangeUnit(out_table, fd, unit);        // 单位转换
            }
            // 删除中间要素
            for (int i = 0; i < case_fields.Count; i++)
            {
                Arcpy.Delect(out_table + i.ToString());
            }
            Arcpy.Delect(out_table + "_total");
        }

        // 属性映射
        public static string AttributeMapper(string in_data, string in_field, string map_field, string map_tabel, bool reverse = false)
        {
            // 获取连接表的2个字段名
            string exl_field01;
            string exl_field02;
            if (reverse)
            {
                exl_field01 = ExcelTool.GetCellFromExcel(map_tabel, 0, 1);
                exl_field02 = ExcelTool.GetCellFromExcel(map_tabel, 0, 0);
            }
            else
            {
                exl_field01 = ExcelTool.GetCellFromExcel(map_tabel, 0, 0);
                exl_field02 = ExcelTool.GetCellFromExcel(map_tabel, 0, 1);
            }

            List<string> fields = new List<string>() { exl_field02 };
            // 连接字段
            Arcpy.JoinField(in_data, in_field, map_tabel, exl_field01, fields);
            // 计算字段
            Arcpy.CalculateField(in_data, map_field, "!" + exl_field02 + "!");
            // 删除多余字段
            Arcpy.DeleteField(in_data, fields);

            return in_data;
        }



    }
}
