using ArcGIS.Core.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Dml.Diagram;
using SharpCompress.Common;
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


namespace CCTool.Scripts.MixApp.StyleMix
{
    /// <summary>
    /// Interaction logic for StylxToFeatureLayer.xaml
    /// </summary>
    public partial class StylxToFeatureLayer : ArcGIS.Desktop.Framework.Controls.ProWindow
    {

        // 工具设置标签
        readonly string toolSet = "StylxToFeatureLayer";
        public StylxToFeatureLayer()
        {
            InitializeComponent();

            combox_type.Items.Add("面符号");
            combox_type.Items.Add("线符号");
            combox_type.Items.Add("点符号");
            combox_type.SelectedIndex = 0;

            // 加载保存的设置
            string featureLayerName = BaseTool.ReadValueFromReg(toolSet, "featureLayerName");
            textFeatureLayerName.Text = (featureLayerName == "") ? "示例图层" : featureLayerName;
        }

        private void combox_stylx_DropDown(object sender, EventArgs e)
        {
            UITool.AddStylxsToComboxPlus(combox_stylx);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string stylxName = combox_stylx.ComboxText();
                string symbolType = combox_type.Text;
                string featureLayerName = textFeatureLayerName.Text;
                string defGDB = Project.Current.DefaultGeodatabasePath;

                string geoType = symbolType switch
                {
                    "面符号" => "Polygon",
                    "线符号" => "Polyline",
                    "点符号" => "Point",
                    _ => "",
                };


                // 判断参数是否选择完全
                if (stylxName == "" || featureLayerName == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 设置保存在本地
                BaseTool.WriteValueToReg(toolSet, "featureLayerName", featureLayerName);

                Close();
                await QueuedTask.Run(() =>
                {
                    string layerPath = $@"{defGDB}\{featureLayerName}";
                    string field = "标记字段";
                    // 创建示例要素类
                    Arcpy.CreateFeatureclass(defGDB, featureLayerName, geoType, "");
                    // 添加一个标记字段，作为唯一值的字段
                    Arcpy.AddField(layerPath, field, "TEXT");
                    // 加载图层
                    MapCtlTool.AddLayerToMap(layerPath);
                    // 获取新加载的图层
                    FeatureLayer featureLayer = featureLayerName.TargetFeatureLayer();
                    // 获取StyleProjectItem
                    StyleProjectItem styleProjectItem = stylxName.TargetStyleProjectItem();
                    // 获取StyleProjectItem
                    List<SymbolStyleItem> symbolStyleItems = StylxTool.GetSymbolStyleItem(styleProjectItem, StyleItemType.PolygonSymbol);

                    // 创建唯一值渲染器配置
                    var uvr = new CIMUniqueValueRenderer
                    {
                        Fields = new[] { field },
                        UseDefaultSymbol = false,
                    };
                    // 预置
                    List<CIMUniqueValueClass> cuvClasses = new List<CIMUniqueValueClass>();

                    // 从样式库中获取相应属性，写到符号系统中
                    foreach (SymbolStyleItem symbolStyleItem in symbolStyleItems)
                    {
                        CIMPolygonSymbol polygonSymbol = symbolStyleItem.Symbol as CIMPolygonSymbol;
                        CIMPointSymbol pointSymbol = symbolStyleItem.Symbol as CIMPointSymbol;
                        CIMLineSymbol polylineSymbol = symbolStyleItem.Symbol as CIMLineSymbol;

                        // 设置
                        CIMSymbolReference cIMSymbolReference = symbolType switch
                        {
                            "面符号" => polygonSymbol.MakeSymbolReference(),
                            "线符号" => polylineSymbol.MakeSymbolReference(),
                            "点符号" => pointSymbol.MakeSymbolReference(),
                            _ => null,
                        };

                        // 设置CIMUniqueValueClass并收集
                        var classA = new CIMUniqueValueClass
                        {
                            Label = symbolStyleItem.Tags,
                            Values = new[]
                            {
                                new CIMUniqueValue { FieldValues = new[] { symbolStyleItem.Name} }
                            },
                            Symbol = cIMSymbolReference,
                            Visible = true
                        };

                        cuvClasses.Add(classA);
                    }

                    // 将分类添加到组
                    var groups = new List<CIMUniqueValueGroup>
                    {
                        new CIMUniqueValueGroup
                        {
                            Classes = cuvClasses.ToArray(),
                        }
                    };

                    uvr.Groups = groups.ToArray();

                    // 应用渲染器
                    featureLayer.SetRenderer(uvr);

                });

                MessageBox.Show($"创建唯一值符号图层完成!");
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147190731";
            UITool.Link2Web(url);
        }


    }
}
