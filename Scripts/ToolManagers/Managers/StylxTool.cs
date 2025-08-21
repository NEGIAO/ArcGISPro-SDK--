using ActiproSoftware.Windows.Extensions;
using ArcGIS.Core.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.ToolManagers.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Managers
{
    public class StylxTool
    {
        // 获取图层的符号样式
        public static Dictionary<string, CIMSymbol> GetFeatureLayerSymbol(FeatureLayer featureLayer)
        {
            //存储唯一值标注与对应符号的字典
            Dictionary<string, CIMSymbol> valueSymbolMap = new Dictionary<string, CIMSymbol>();

            // 获取唯一值渲染器
            CIMUniqueValueRenderer renderer = featureLayer.GetRenderer() as CIMUniqueValueRenderer;

            //遍历渲染器的所有分组及分类
            foreach (var group in renderer.Groups)
            {
                foreach (CIMUniqueValueClass uvc in group.Classes)
                {
                    // 提取标注（Label）和对应的符号引用（Symbol）
                    string label = uvc.Label;
                    CIMSymbol symbol = uvc.Symbol.Symbol;

                    // 添加到字典（假设标注唯一）
                    if (!valueSymbolMap.ContainsKey(label))
                        valueSymbolMap.Add(label, symbol);
                }
            }

            return valueSymbolMap;
        }

        // 获取图层的符号样式【复杂】
        public static Dictionary<string, StylxAtt> GetFeatureLayerSymbolDict(FeatureLayer featureLayer)
        {
            //存储唯一值标注与对应符号的字典
            Dictionary<string, StylxAtt> StylxAtts = new Dictionary<string, StylxAtt>();

            // 获取唯一值渲染器
            CIMUniqueValueRenderer renderer = featureLayer.GetRenderer() as CIMUniqueValueRenderer;

            //遍历渲染器的所有分组及分类
            foreach (var group in renderer.Groups)
            {
                foreach (CIMUniqueValueClass uvc in group.Classes)
                {
                    StylxAtt stylxAtt = new StylxAtt();
                    // 提取value，标注（Label）和对应的符号引用（Symbol）
                    string label = uvc.Label;
                    string value = uvc.Values.FirstOrDefault().FieldValues.FirstOrDefault();
                    CIMSymbol symbol = uvc.Symbol.Symbol;

                    // 赋值
                    stylxAtt.Name = value;
                    stylxAtt.Tag = label;
                    stylxAtt.Symbol = symbol;

                    // 添加到字典（假设标注唯一）
                    if (!StylxAtts.ContainsKey(value))
                        StylxAtts.Add(value, stylxAtt);
                }
            }

            return StylxAtts;
        }


        // 获取Stylx的所有StyleItem
        public static List<StyleItem> GetStyleItem(StyleProjectItem styleProjectItem, StyleItemType styleItemType = StyleItemType.Unknown)
        {
            // 收集所有StyleItem
            List<StyleItem> allStyleItems = new List<StyleItem>();
            // 全部类型
            if (styleItemType == StyleItemType.Unknown)
            {
                // 获取所有符号类项目（指定支持的Symbol类型）
                var symbolTypes = new[]
                {
                StyleItemType.PointSymbol,
                StyleItemType.LineSymbol,
                StyleItemType.PolygonSymbol,
                StyleItemType.TextSymbol,
                StyleItemType.NorthArrow,
                StyleItemType.Grid,
            };

                foreach (var type in symbolTypes)
                {
                    // 搜索指定Symbol类型的所有项（空字符串表示全部）
                    var symbols = styleProjectItem.SearchSymbols(type, "");
                    allStyleItems.AddRange(symbols);
                }

                // 获取所有颜色梯度
                var colorRamps = styleProjectItem.SearchColorRamps("");
                allStyleItems.AddRange(colorRamps);
            }
            // 单独类型
            else
            {
                // 搜索指定Symbol类型的所有项（空字符串表示全部）
                var symbols = styleProjectItem.SearchSymbols(styleItemType, "");
                allStyleItems.AddRange(symbols);
            }

            return allStyleItems;
        }


        // 获取Stylx的所有Item
        public static List<SymbolStyleItem> GetSymbolStyleItem(StyleProjectItem styleProjectItem, StyleItemType styleItemType = StyleItemType.Unknown)
        {
            // 收集所有StyleItem
            List<SymbolStyleItem> allStyleItems = new List<SymbolStyleItem>();

            // 全部类型
            if (styleItemType == StyleItemType.Unknown)
            {
                // 获取所有符号类项目（指定支持的Symbol类型）
                var symbolTypes = new[]
                {
                    StyleItemType.PointSymbol,
                    StyleItemType.LineSymbol,
                    StyleItemType.PolygonSymbol,
                    StyleItemType.TextSymbol,
                    StyleItemType.NorthArrow,
                    StyleItemType.Grid,
                };

                foreach (var type in symbolTypes)
                {
                    // 搜索指定Symbol类型的所有项（空字符串表示全部）
                    var symbols = styleProjectItem.SearchSymbols(type, "");
                    allStyleItems.AddRange(symbols);
                }

            }
            // 单独类型
            else
            {
                // 搜索指定Symbol类型的所有项（空字符串表示全部）
                var symbols = styleProjectItem.SearchSymbols(styleItemType, "");
                allStyleItems.AddRange(symbols);
            }
            return allStyleItems;
        }

        // 获取Stylx的所有Item的Name
        public static List<string> GetStyleItemNames(StyleProjectItem styleProjectItem, StyleItemType styleItemType = StyleItemType.Unknown)
        {
            List<string> itemNames = new List<string>();

            // 收集所有StyleItem
            List<StyleItem> allStyleItems = GetStyleItem(styleProjectItem, styleItemType);
            // 收集Name
            foreach (var styleItem in allStyleItems)
            {
                itemNames.Add(styleItem.Name);
            }

            return itemNames;
        }

        // 创建一个新样式
        public static void CreateStylx(string stylxName)
        {
            string defFolder = Project.Current.HomeFolderPath;
            // 新建的路径
            string defPath = $@"{defFolder}\{stylxName}.stylx";

            //获取当前工程中的所有tylx
            var ProjectStyles = Project.Current.GetItems<StyleProjectItem>();
            //根据名字找出指定的tylx
            StyleProjectItem style = ProjectStyles.FirstOrDefault(x => x.Name == stylxName);
            
            // 如果当前地图存在该stylx，就先删除
            if (style is not null)
            {
                // stylx文件原始路径
                string stylePath = style.Path;
                StyleHelper.RemoveStyle(Project.Current, stylePath);
                File.Delete(stylePath);
            }

            StyleHelper.CreateStyle(Project.Current, defPath);
        }

        // 清空Stylx内容
        public static void ClearStyleItem(StyleProjectItem styleProjectItem)
        {
            List<StyleItem> allStyleItems = GetStyleItem(styleProjectItem);

            // 清除
            foreach (StyleItem styleItem in allStyleItems)
            {
                styleProjectItem.RemoveItem(styleItem);
            }
        }

        // 把Dictionary<string, CIMSymbol>写入样式库文件
        public static void WriteStylxItem(StyleProjectItem styleProjectItem, Dictionary<string, StylxAtt> valueSymbolMap)
        {
            // stylx原itemName
            List<string> originItemNames = GetStyleItemNames(styleProjectItem);

            // 写入
            foreach (var item in valueSymbolMap)
            {
                // 如果已经同名item就清除原来的
                if (originItemNames.Contains(item.Key))
                {
                    styleProjectItem.RemoveItemByName(item.Key);
                }

                StylxAtt stylxAtt = item.Value;

                // 创建符号样式项
                var symbolItem = new SymbolStyleItem()
                {
                    Name = stylxAtt.Name,  //符号名称
                    Category = "", //分类目录
                    Tags = stylxAtt.Tag, //标签
                    Key = stylxAtt.Name,  //唯一标识（可选）
                    Symbol = stylxAtt.Symbol,
                };

                // 将符号添加到样式库
                styleProjectItem.AddItem(symbolItem);
            }
        }

        // Stylx的Value和Tag对调
        public static void ChangeValueAndTag(StyleProjectItem styleProjectItem)
        {
            List<SymbolStyleItem> allStyleItems = GetSymbolStyleItem(styleProjectItem);

            // 替换
            foreach (SymbolStyleItem styleItem in allStyleItems)
            {
                // 创建符号样式项
                SymbolStyleItem newItem = new SymbolStyleItem()
                {
                    Name = styleItem.Tags,  //符号名称
                    Category = styleItem.Category, //分类目录
                    Tags = styleItem.Name, //标签
                    Symbol = styleItem.Symbol,
                };
                StyleHelper.RemoveItem(styleProjectItem, styleItem);
                StyleHelper.AddItem(styleProjectItem, newItem);
            }
        }


    }

    public class StylxAtt
    {
        public string Name { get; set; }
        public string Category { get; set; }
        public string Tag { get; set; }
        public string Key { get; set; }
        public CIMSymbol Symbol { get; set; }
        
    }

}


