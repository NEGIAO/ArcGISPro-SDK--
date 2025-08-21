using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Library
{
    public class GlobalClass
    {

    }

    // 点属性
    public class PointAtt
    {
        public string Name { get; set; }
        public string Des { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
    }


    // 要素类属性
    public class FeatureClassAtt
    {
        public string Name { get; set; }                    // 要素名
        public string AliasName { get; set; }              // 要素别名
        public IReadOnlyList<FieldDescription> FieldDescriptions { get; set; }          // 字段列表
        public SpatialReference SpatialReference { get; set; }                  // 坐标系
        public long FeatureCount { get; set; }                       // 要素数量
        public string OIDField { get; set; }                             // objectID字段名
        public GeometryType GeometryType { get; set; }           // 要素类型
        public bool HasZ { get; set; }                           // 是否有Z值
        public bool HasM { get; set; }                          // 是否有M值
    }

    // 字段属性
    public class FieldAtt
    {
        public string Name { get; set; }                    // 字段名
        public string AliasName { get; set; }              // 字段别名
        public FieldType Type { get; set; }          // 字段类型
        public int Length { get; set; }                  // 字段长度
    }
}
