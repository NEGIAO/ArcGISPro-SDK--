using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.ToolManagers.Extensions;
using NPOI.OpenXmlFormats.Vml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Managers
{
    public class CheckTool
    {
        // 检查路径是否是gdb
        public static string CheckGDBPath(string path)
        {
            string result = "";
            // 是否包含.gdb
            if (path.Contains(".gdb") == false)
            {
                result += $"非GDB路径：【{path}】\r";
            }

            return result;
        }

        // 检查gdb要素是否是以数字开关
        public static string CheckGDBIsNumeric(string fcPath)
        {
            string result = "";
            // 获取目标数据库和点要素名
            string fcName = fcPath[(fcPath.LastIndexOf("\\") + 1)..];

            // 判断要素名是不是以数字开头
            bool isNum = fcName.IsNumeric();
            if (isNum)
            {
                result += $"GDB要素【{fcName}】不能以数字开头！\r";
            }

            return result;
        }

        // 检查gdb下的要素是否合规
        public static string CheckGDBFeature(string fcPath)
        {
            string result = "";
            // 检查路径是否是gdb
            string re1 = CheckGDBPath(fcPath);
            result += re1;

            // 如果路径是gdb
            if (re1 == "")
            {
                // 检查GDB路径是否存在
                string gdbPath = fcPath[..(fcPath.LastIndexOf(".gdb") + 4)];
                bool isHaveGDB = Directory.Exists(gdbPath);
                if (!isHaveGDB)
                {
                    string re2 = $"GDB路径【{gdbPath}】不存在！\r";
                    result += re2;
                }
            }

            // 检查gdb要素是否是以数字开关
            string re3 = CheckGDBIsNumeric(fcPath);
            result += re3;

            return result;
        }


        // 判断是否有Z值
        public static string CheckHasZ(string target_fc)
        {
            string result = "";

            FeatureClass featureClass = target_fc.TargetFeatureClass();

            // 遍历面要素类中的所有要素
            RowCursor cursor = featureClass.Search();
            while (cursor.MoveNext())
            {
                using var feature = cursor.Current as Feature;
                // 获取要素的几何
                Polygon geometry = feature.GetShape() as Polygon;

                if (geometry.HasZ)
                {
                    result = "输入的要素有Z值，请先清除掉！\r";
                    break;
                }
            }
            return result;
        }


        // 检查要素图层的数据类型是否正确
        public static string CheckFeatureClassType(string lyName, string fcType)
        {
            string result = "";

            // FeatureLayer
            FeatureLayer init_featurelayer = lyName.TargetFeatureLayer();

            if (init_featurelayer is not null)
            {
                // 获取要素图层的要素类（FeatureClass）
                FeatureClass featureClass = init_featurelayer.GetFeatureClass();

                // 获取要素类的几何类型
                string featureClassType = featureClass.GetDefinition().GetShapeType().ToString();

                // 判断
                if (featureClassType != fcType)
                {
                    result += $"【{lyName}】的要素类型不是【{fcType}】\r";
                }
            }

            return result;
        }

        // 检查目标是否包含相应字段
        public static string IsHaveFieldInTarget(string targetPath, string fieldName)
        {
            string result = "";
            // 获取属性表
            Table table = targetPath.TargetTable();
            // 获取字段
            var fields = table.GetDefinition().GetFields();
            // 收集字段名
            List<string> tableFields = new List<string>();
            foreach (Field field in fields)
            {
                tableFields.Add(field.Name);
            }

            // 提取错误信息
            if (!tableFields.Contains(fieldName))
            {
                result += $"【{targetPath}】中缺少【{fieldName}】字段。\r";
            }

            return result;
        }

        // 检查目标是否包含相应字段【多个字段】
        public static string IsHaveFieldInTarget(string targetPath, List<string> fieldNames)
        {
            string result = "";
            // 获取属性表
            Table table = targetPath.TargetTable();
            // 获取字段
            var fields = table.GetDefinition().GetFields();
            // 收集字段名
            List<string> tableFields = new List<string>();
            foreach (Field field in fields)
            {
                tableFields.Add(field.Name);
            }

            // 提取错误信息
            foreach (string fieldName in fieldNames)
            {
                if (!tableFields.Contains(fieldName))
                {
                    result += $"【{targetPath}】中缺少【{fieldName}】字段。\r";
                }
            }

            return result;
        }


        // 检查图层中是否包含相应字段
        public static string IsHaveFieldInLayer(string lyName, string fieldName)
        {
            string result = "";
            List<string> str_fields = new List<string>();
            List<FieldDescription> fields = new List<FieldDescription>();

            var init_featurelayer = lyName.TargetFeatureLayer();
            var init_table = lyName.TargetStandaloneTable();

            // 判断当前选择的是要素图层还是独立表
            if (init_table is not null)
            {
                fields = init_table.GetFieldDescriptions();
            }
            else if (init_featurelayer is not null)
            {
                fields = init_featurelayer.GetFieldDescriptions();
            }
            // 生成字段列表
            foreach (var item in fields)
            {
                str_fields.Add(item.Name.ToLower());
            }
            // 提取错误信息
            if (!str_fields.Contains(fieldName.ToLower()))
            {
                result += $"【{lyName}】中缺少【{fieldName}】字段";
            }

            return result;
        }

        // 检查图层中是否包含相应字段【多个字段】
        public static string IsHaveFieldInLayer(string lyName, List<string> fieldName)
        {
            string result = "";
            List<string> str_fields = new List<string>();
            List<FieldDescription> fields = new List<FieldDescription>();

            var init_featurelayer = lyName.TargetFeatureLayer();
            var init_table = lyName.TargetStandaloneTable();

            // 判断当前选择的是要素图层还是独立表
            if (init_table is not null)
            {
                fields = init_table.GetFieldDescriptions();
            }
            else if (init_featurelayer is not null)
            {
                fields = init_featurelayer.GetFieldDescriptions();
            }
            // 生成字段列表
            foreach (var item in fields)
            {
                str_fields.Add(item.Name);
            }
            // 提取错误信息
            foreach (var item in fieldName)
            {
                if (!str_fields.Contains(item))
                {
                    result += $"【{lyName}】中缺少【{item}】字段\r";
                }
            }

            return result;
        }

        // 检查是否正常提取Excel
        public static string CheckExcelPick()
        {
            string result = "";
            // 复制一个模板
            string def_path = Project.Current.HomeFolderPath;
            string excel_mapper = $@"{def_path}\测试模板.xlsx";
            DirTool.CopyResourceFile(@$"CCTool.Data.Excel.管控边界表.xlsx", excel_mapper);
            Dictionary<string, string> MapDict = ExcelTool.GetDictFromExcel(excel_mapper);
            if (MapDict.Count == 0)
            {
                result = "无法正常读取Excel，请安装相关补丁！";
            }
            File.Delete(excel_mapper);
            return result;
        }


        // 检查字段值是否符合要求
        public static string CheckFieldValue(string lyName, string check_field, List<string> checkStringList)
        {
            string result = "";

            // 判断是否有这个字段
            string result_isHaveField = IsHaveFieldInLayer(lyName, check_field);
            result += result_isHaveField;

            // 获取ID字段
            string IDField = lyName.TargetIDFieldName();

            // 如果没有字段缺失，刚进行下一步判断
            if (!result_isHaveField.Contains("】中缺少【"))
            {
                // 获取Table
                Table table = lyName.TargetTable();
                // 逐行找出错误
                using RowCursor rowCursor = table.Search();
                while (rowCursor.MoveNext())
                {
                    using Row row = rowCursor.Current;
                    // 获取value
                    var va = row[check_field];
                    if (va == null)
                    {
                        result += $"({IDField}:{row[IDField]})：【{lyName}】中的【{check_field}】字段存在空值\r";
                    }
                    else
                    {
                        string value = va.ToString();
                        if (!checkStringList.Contains(value))
                        {
                            result += $"({IDField}:{row[IDField]})：【{lyName}】中的【{check_field}】字段存在不符合要求的字段值【{value}】\r";
                        }
                    }
                }
            }
            return result;
        }

        // 检查字段值是否有空值
        public static string CheckFieldValueEmpty(string lyName, string check_field, string sql = "")
        {
            string result = "";
            List<string> str_fields = new List<string>();
            List<FieldDescription> fields = new List<FieldDescription>();

            // 判断是否有这个字段
            string result_isHaveField = IsHaveFieldInLayer(lyName, check_field);
            result += result_isHaveField;

            // 获取ID字段
            string IDField = lyName.TargetIDFieldName();

            if (!result_isHaveField.Contains("】中缺少【"))
            {
                // 判断当前选择的是要素图层还是独立表
                Table table = lyName.TargetTable();

                var queryFilter = new QueryFilter();
                queryFilter.WhereClause = sql;
                // 逐行找出错误
                using RowCursor rowCursor = table.Search(queryFilter);
                while (rowCursor.MoveNext())
                {
                    using Row row = rowCursor.Current;
                    // 获取value
                    var va = row[check_field];
                    if (va == null)
                    {
                        result += $"({IDField}:{row[IDField]})：【{lyName}】中的【{check_field}】字段存在空值\r";
                    }
                }
            }
            return result;
        }

        // 检查字段值是否是空字符串和空值
        public static string CheckFieldValueSpace(string lyName, string check_field, string sql = "")
        {
            string result = "";
            List<string> str_fields = new List<string>();
            List<FieldDescription> fields = new List<FieldDescription>();

            // 判断是否有这个字段
            string result_isHaveField = IsHaveFieldInLayer(lyName, check_field);
            result += result_isHaveField;

            // 获取ID字段
            string IDField = lyName.TargetIDFieldName();

            if (!result_isHaveField.Contains("】中缺少【"))
            {
                // 判断当前选择的是要素图层还是独立表
                Table table = lyName.TargetTable();

                var queryFilter = new QueryFilter();
                queryFilter.WhereClause = sql;
                // 逐行找出错误
                using RowCursor rowCursor = table.Search(queryFilter);
                while (rowCursor.MoveNext())
                {
                    using Row row = rowCursor.Current;
                    // 获取value
                    var va = row[check_field];
                    if (va == null)
                    {
                        result += $"({IDField}:{row[IDField]})：【{lyName}】中的【{check_field}】字段存在空值\r";
                    }
                    else
                    {
                        if (va.ToString().Replace(" ", "") == "")
                        {
                            result += $"({IDField}:{row[IDField]})：【{lyName}】中的【{check_field}】字段存在空字符串\r";
                        }
                        else if (va.ToString() == "0")
                        {
                            result += $"({IDField}:{row[IDField]})：【{lyName}】中的【{check_field}】字段值为0\r";
                        }
                    }
                }
            }
            return result;
        }

        // 检查字段值是否有空值【多个字段】
        public static string CheckFieldValueEmpty(string lyName, List<string> check_field)
        {
            string result = "";
            List<string> str_fields = new List<string>();
            List<FieldDescription> fields = new List<FieldDescription>();

            // 判断是否有这个字段
            string result_isHaveField = IsHaveFieldInLayer(lyName, check_field);
            result += result_isHaveField;

            // 获取ID字段
            string IDField = lyName.TargetIDFieldName();

            foreach (string fd in check_field)
            {
                if (!result_isHaveField.Contains(fd))
                {
                    // 获取Table
                    Table table = lyName.TargetTable();

                    // 逐行找出错误
                    using RowCursor rowCursor = table.Search();
                    while (rowCursor.MoveNext())
                    {
                        using Row row = rowCursor.Current;
                        // 获取value
                        var va = row[fd];
                        if (va == null)
                        {
                            result += $"({IDField}:{row[IDField]})：【{lyName}】中的【{fd}】字段存在空值\r";
                        }
                    }
                }
            }

            return result;
        }

        // 检查字段值是否有空字符串和空值【多个字段】
        public static string CheckFieldValueSpace(string lyName, List<string> check_field)
        {
            string result = "";
            List<string> str_fields = new List<string>();
            List<FieldDescription> fields = new List<FieldDescription>();

            // 判断是否有这个字段
            string result_isHaveField = IsHaveFieldInLayer(lyName, check_field);
            result += result_isHaveField;

            // 获取ID字段
            string IDField = lyName.TargetIDFieldName();

            foreach (string fd in check_field)
            {
                if (!result_isHaveField.Contains(fd))
                {
                    // 获取Table
                    Table table = lyName.TargetTable();

                    // 逐行找出错误
                    using RowCursor rowCursor = table.Search();
                    while (rowCursor.MoveNext())
                    {
                        using Row row = rowCursor.Current;
                        // 获取value
                        var va = row[fd];
                        if (va == null)
                        {
                            result += $"({IDField}:{row[IDField]})：【{lyName}】中的【{fd}】字段存在空值\r";
                        }
                        else
                        {
                            if (va.ToString().Replace(" ", "") == "")
                            {
                                result += $"({IDField}:{row[IDField]})：【{lyName}】中的【{check_field}】字段存在空字符串\r";
                            }
                        }
                    }
                }
            }

            return result;
        }

        // 检查要素是否存在多部件和空洞
        public static string CheckMultiPart(string lyName)
        {
            string result = "";

            // 获取ID字段
            string IDField = lyName.TargetIDFieldName();

            // 获取FeatureClass
            FeatureClass featureClass = lyName.TargetFeatureClass();

            // 逐行找出错误
            using RowCursor rowCursor = featureClass.Search();
            while (rowCursor.MoveNext())
            {
                using Feature feature = rowCursor.Current as Feature;

                // 获取要素的几何
                Polygon geometry = feature.GetShape() as Polygon;
                // 获取部件数
                int partCount = geometry.PartCount;

                if (partCount > 1)
                {
                    result += $"({IDField}:{feature[IDField]})：【{lyName}】中存在多部件或空洞\r";
                }
            }

            return result;
        }

        // 检查2个图层的坐标系是否一致
        public static string CheckSpatialReference(string featureLayer1, string featureLayer2)
        {
            string result = "";

            // 获取坐标系的名称
            string srName1 = featureLayer1.TargetFeatureLayer().GetSpatialReference().Name;

            string srName2 = featureLayer2.TargetFeatureLayer().GetSpatialReference().Name;


            if (srName1 != srName2)
            {
                result += $"【{featureLayer1}】和【{featureLayer2}】的坐标系不一致。\r";
            }

            return result;
        }

        // 检查文件夹路径是否存在
        public static string CheckFolderExists(string folder)
        {
            string result = "";
            bool isExist = Directory.Exists(folder);
            if (!isExist)
            {
                result += $"文件夹路径不存在：{folder}\r";
            }
            return result;
        }
    }
}
