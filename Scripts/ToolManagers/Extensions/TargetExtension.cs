using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Extensions
{
    public static class TargetExtension
    {

        // 通过路径或图层名获取StyleProjectItem
        public static StyleProjectItem TargetStyleProjectItem(this string styleName)
        {
            if (styleName == null)
            {
                return null;
            }

            //获取当前工程中的所有tylx
            var ProjectStyles = Project.Current.GetItems<StyleProjectItem>();
            //根据名字找出指定的stylx
            StyleProjectItem styleProjectItem = ProjectStyles.FirstOrDefault(x => x.Name == styleName);

            return styleProjectItem;
        }
        
        // 通过路径或图层名获取输入的工作空间
        public static string TargetWorkSpace(this string sourcePath)
        {
            if (sourcePath == null)
            {
                return null;
            }

            // 兼容两种符号
            string inputData = sourcePath.Replace(@"/", @"\");
            // 如果是GDB文件路径
            if (inputData.Contains(".gdb"))
            {
                // 获取最后一个".gdb"的位置
                int index = inputData.LastIndexOf(@".gdb");
                // 获取gdb文件名
                string gdbPath = inputData[..(index + 4)];
                // 返回gdb路径
                return gdbPath;
            }
            // 如果是SHP文件路径
            else if (inputData.Contains(".shp"))
            {
                // 获取最后一个"\"的位置
                int index = inputData.LastIndexOf(@"\");
                // 获取要素名
                string shpPath = inputData[..index];
                // 返回shp文件路径
                return shpPath;
            }
            // 如果是图层名
            else
            {
                // 获取图层数据的完整路径
                string fullPath = inputData.TargetLayerPath();
                // 递归返回工作空间
                return fullPath.TargetWorkSpace();
            }
        }

        // 通过路径或图层名获取输入的要素名称
        public static string TargetFcName(this string sourcePath)
        {
            if (sourcePath == null)
            {
                return null;
            }

            // 兼容两种符号
            string inputData = sourcePath.Replace(@"/", @"\");
            // 如果是GDB文件路径
            if (inputData.Contains(".gdb"))
            {
                // 获取最后一个"\"的位置
                int index = inputData.LastIndexOf(@"\");
                // 获取要素名
                string targetName = inputData[(index + 1)..];
                // 返回
                return targetName;
            }
            // 如果是SHP文件路径
            else if (inputData.Contains(".shp"))
            {
                // 获取最后一个"\"的位置
                int index = inputData.LastIndexOf(@"\");
                // 获取要素名
                string targetName = inputData[(index + 1)..];
                // 返回
                return targetName;
            }
            // 如果是图层名
            else
            {
                // 获取图层数据的完整路径
                string fullPath = inputData.TargetLayerPath();
                // 递归返回要素名
                return fullPath.TargetFcName();
            }
        }

        // 通过图层名获取数据的完整路径
        public static string TargetLayerPath(this string layerName)
        {
            Map map = MapView.Active.Map;
            Dictionary<FeatureLayer, string> dic_ly = map.AllFeatureLayersDic();
            Dictionary<StandaloneTable, string> dic_table = map.AllStandaloneTablesDic();

            // 获取完整路径
            string targetPath = "";
            // 如果是图层
            if (dic_ly.ContainsValue(layerName))
            {
                foreach (var item in dic_ly)
                {
                    if (item.Value == layerName)
                    {
                        var p = item.Key.GetPath();
                        if (p != null)
                        {
                            targetPath = p.ToString().Replace("file:///", "").Replace("/", @"\");
                        }
                    }
                }
            }
            // 如果是独立表
            else if (dic_table.ContainsValue(layerName))
            {
                foreach (var item in dic_table)
                {
                    if (item.Value == layerName)
                    {
                        var p = item.Key.GetPath();
                        if (p != null)
                        {
                            targetPath = p.ToString().Replace("file:///", "").Replace("/", @"\");
                        }
                    }
                }
            }

            // shp的情况，需要加上.shp
            if (targetPath != "" && !targetPath.Contains(".gdb"))
            {
                if (!targetPath.ToLower().Contains(".dwg") && !targetPath.ToLower().Contains(".dxf"))    // 再排除cad的情况
                {
                    targetPath += ".shp";
                }
            }
            // 返回完整路径
            return targetPath;
        }

        // 通过图层获取数据的完整路径
        public static string TargetLayerPath(this FeatureLayer featureLayer)
        {
            string lp = "";
            string layerPath;
            try
            {
                layerPath = featureLayer.GetPath().ToString();
            }
            catch (Exception)
            {

                layerPath = "";
            }

            
            if (layerPath is not null || layerPath == "")
            {
                lp = layerPath.ToString().Replace("file:///", "").Replace("/", @"\");
            }

            // shp的情况，需要加上.shp
            if (lp != "" && !lp.Contains(".gdb"))
            {
                if (!lp.ToLower().Contains(".dwg") && !lp.ToLower().Contains(".dxf"))   // 还要排除CAD数据
                {
                    lp += ".shp";
                }
            }
            // 返回完整路径
            return lp;
        }

        // 通过路径或图层名获取目标要素FeatureClass
        public static FeatureClass TargetFeatureClass(this string filePath)
        {
            // 获取目录的路径和名称
            string targetPath = filePath.TargetWorkSpace();
            string targetName = filePath.TargetFcName();

            // 如果是GDB数据
            if (filePath.Contains(".gdb"))
            {
                using Geodatabase geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(targetPath)));
                FeatureClass featureClass = geodatabase.OpenDataset<FeatureClass>(targetName);
                return featureClass;
            }
            // 如果是SHP数据
            else if (filePath.Contains(".shp"))
            {
                FileSystemConnectionPath connectionPath = new FileSystemConnectionPath(new Uri(targetPath), FileSystemDatastoreType.Shapefile);
                using FileSystemDatastore shapefile = new FileSystemDatastore(connectionPath);
                FeatureClass featureClass = shapefile.OpenDataset<FeatureClass>(targetName);
                return featureClass;
            }
            else
            {
                // 获取图层的完整路径
                string layerSourcePath = filePath.TargetLayerPath();
                FeatureClass featureClass = layerSourcePath.TargetFeatureClass();
                return featureClass;
            }
        }

        // 通过路径或图层名获取目标要素FeatureClass
        public static FeatureClassDescription TargetFeatureClassDescription(this string filePath)
        {
            // 获取目录的路径和名称
            string targetPath = filePath.TargetWorkSpace();
            string targetName = filePath.TargetFcName();
            // 如果是GDB数据
            if (filePath.Contains(".gdb"))
            {
                using Geodatabase geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(targetPath)));
                // 获取待修改要素的【FeatureClassDefinition】
                FeatureClassDefinition featureClassDefinition = geodatabase.GetDefinition<FeatureClassDefinition>(targetName);
                // 获取【FeatureClassDescription】
                FeatureClassDescription featureClassDescription = new FeatureClassDescription(featureClassDefinition);
                return featureClassDescription;
            }
            // 如果是SHP数据
            else if (filePath.Contains(".shp"))
            {
                FileSystemConnectionPath connectionPath = new FileSystemConnectionPath(new Uri(targetPath), FileSystemDatastoreType.Shapefile);
                using FileSystemDatastore shapefile = new FileSystemDatastore(connectionPath);
                // 获取待修改要素的【FeatureClassDefinition】
                FeatureClassDefinition featureClassDefinition = shapefile.GetDefinition<FeatureClassDefinition>(targetName);
                // 获取【FeatureClassDescription】
                FeatureClassDescription featureClassDescription = new FeatureClassDescription(featureClassDefinition);
                return featureClassDescription;
            }
            else
            {
                // 获取图层的完整路径
                string layerSourcePath = filePath.TargetLayerPath();
                FeatureClassDescription featureClassDescription = layerSourcePath.TargetFeatureClassDescription();
                return featureClassDescription;
            }
        }

        // 通过图层名获取目标要素的FeatureLayer
        public static FeatureLayer TargetFeatureLayer(this string layerName)
        {
            Object ob = layerName.GetLayerFromFullName();
            if (ob is FeatureLayer)
            {
                return (FeatureLayer)ob;
            }
            else { return null; }
        }

        // 通过图层名获取目标要素的StandaloneTable
        public static StandaloneTable TargetStandaloneTable(this string layerName)
        {
            Object ob = layerName.GetLayerFromFullName();
            if (ob is StandaloneTable)
            {
                return (StandaloneTable)ob;
            }
            else { return null; }
        }

        // 通过路径或图层名获取目标要素的属性表
        public static Table TargetTable(this string filePath)
        {
            // 获取目录的路径和名称
            string targetPath = filePath.TargetWorkSpace();
            string targetName = filePath.TargetFcName();
            // 如果是GDB数据
            if (filePath.Contains(".gdb"))
            {
                using Geodatabase geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(targetPath)));
                Table table = geodatabase.OpenDataset<Table>(targetName);
                return table;
            }
            // 如果是SHP数据
            else if (filePath.Contains(".shp"))
            {
                FileSystemConnectionPath connectionPath = new FileSystemConnectionPath(new Uri(targetPath), FileSystemDatastoreType.Shapefile);
                using FileSystemDatastore shapefile = new FileSystemDatastore(connectionPath);
                Table table = shapefile.OpenDataset<Table>(targetName);
                return table;
            }
            else
            {
                // 获取图层的完整路径
                string layerSourcePath = filePath.TargetLayerPath();
                Table table = layerSourcePath.TargetTable();
                return table;
            }
        }

        // 通过路径或图层名获取目标要素的TableDefinition
        public static TableDefinition TargetTableDefinition(this string filePath)
        {
            // 获取目录的路径和名称
            string targetPath = filePath.TargetWorkSpace();
            string targetName = filePath.TargetFcName();
            // 如果是GDB数据
            if (filePath.Contains(".gdb"))
            {
                using Geodatabase geodatabase = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(targetPath)));
                TableDefinition table = geodatabase.GetDefinition<TableDefinition>(targetName);
                return table;
            }
            // 如果是SHP数据
            else if (filePath.Contains(".shp"))
            {
                FileSystemConnectionPath connectionPath = new FileSystemConnectionPath(new Uri(targetPath), FileSystemDatastoreType.Shapefile);
                using FileSystemDatastore shapefile = new FileSystemDatastore(connectionPath);
                TableDefinition table = shapefile.GetDefinition<TableDefinition>(targetName);
                return table;
            }
            else
            {
                // 获取图层的完整路径
                string layerSourcePath = filePath.TargetLayerPath();
                TableDefinition table = layerSourcePath.TargetTableDefinition();
                return table;
            }
        }

        // 获取要素类的坐标系
        public static SpatialReference TargetSpatialReference(this string fcPath)
        {
            // 获取FeatureClass
            FeatureClass featureClass = fcPath.TargetFeatureClass();
            SpatialReference sr = featureClass.GetDefinition().GetSpatialReference();

            // 返回坐标系
            return sr;
        }

        // 看一下是否当前有选择图斑，如果没有选择图斑，就全图斑处理，如果有选择图斑，就按选择图斑处理
        public static RowCursor TargetSelectCursor(this FeatureLayer featureLayer)
        {
            // 看一下是否当前有选择图斑
            RowCursor cursor;
            // 如果没有选择图斑，就全图斑处理
            if (featureLayer.SelectionCount == 0)
            {
                cursor = featureLayer.Search();
            }
            // 如果有选择图斑，就按选择图斑处理
            else
            {
                cursor = featureLayer.GetSelection().Search();
            }
            // 返回
            return cursor;
        }

        // 看一下是否当前有选择图斑，如果没有选择图斑，就全图斑处理，如果有选择图斑，就按选择图斑处理
        public static RowCursor TargetSelectCursor(this StandaloneTable sandaloneTable)
        {
            // 看一下是否当前有选择图斑
            RowCursor cursor;
            // 如果没有选择图斑，就全图斑处理
            if (sandaloneTable.SelectionCount == 0)
            {
                cursor = sandaloneTable.Search();
            }
            // 如果有选择图斑，就按选择图斑处理
            else
            {
                cursor = sandaloneTable.GetSelection().Search();
            }
            // 返回
            return cursor;
        }

        // 从路径或图层中获取要素类型
        public static GeometryType TargetGeoType(this object layerName)
        {
            GeometryType geoType = GeometryType.Unknown;
            // 获取FeatureLayer
            FeatureClass featureClass;
            if (layerName is string @string)
            {
                featureClass = @string.TargetFeatureClass();
            }
            else if (layerName is FeatureLayer layer)
            {
                featureClass = layer.GetFeatureClass();
            }
            else
            {
                featureClass = null;
            }
            // 定位到要素
            if (featureClass != null)
            {
                geoType = featureClass.GetDefinition().GetShapeType();
            }
            return geoType;
        }

        // 从路径或图层中获取对象ID字段
        public static string TargetIDFieldName(this object layerName)
        {
            //  获取不可编辑的字段
            List<Field> fields = GisTool.GetFieldsFromTarget(layerName, "notEdit");
            string IDField = "";
            foreach (var field in fields)
            {
                if (field.FieldType == FieldType.OID)
                {
                    IDField = field.Name;
                }
            }
            return IDField;
        }

        // 从路径或图层中获取指定字段，指定行的value
        public static string TargetCellValue(this string inputData, string inputField, string sql = "")
        {
            string value = "";

            // sql预处理
            if (sql == "")
            {
                var oidField = inputData.TargetIDFieldName();
                sql = $"{oidField}=1";
            }

            // 获取Table
            Table table = inputData.TargetTable();
            // 设定筛选语句
            var queryFilter = new QueryFilter();
            queryFilter.WhereClause = sql;
            // 逐行搜索
            using RowCursor rowCursor = table.Search(queryFilter, false);
            while (rowCursor.MoveNext())
            {
                using Row row = rowCursor.Current;
                // 获取value
                var va = row[inputField];
                if (va is not null)
                {
                    value = va.ToString();
                }
            }
            return value;
        }
    }
}
