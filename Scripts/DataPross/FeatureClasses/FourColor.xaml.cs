using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Words.Fields;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    /// <summary>
    /// Interaction logic for FourColor.xaml
    /// </summary>
    public partial class FourColor : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public FourColor()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "四色表达";

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            string fc = combox_fc.ComboxText();
            UITool.AddTextFieldsToComboxPlus(fc, combox_field);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string field = combox_field.ComboxText();

                // 判断参数是否选择完全
                if (fc_path == "" || field == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("字段查询，并填写颜色名");
                    
                    // 获取原始图层和标识图层
                    FeatureLayer originFeatureLayer = fc_path.TargetFeatureLayer();
                    // 获取原始图层和标识图层的要素类
                    FeatureClass originFeatureClass = fc_path.TargetFeatureClass();

                    // 获取目标图层和源图层的要素游标
                    using RowCursor originCursor = originFeatureClass.Search();
                    // 遍历源图层的要素
                    while (originCursor.MoveNext())
                    {
                        // 预设4种颜色名称
                        List<string> colors = new List<string>()
                            {
                                "颜色01","颜色02","颜色03","颜色04","颜色05","颜色06","颜色07","颜色08"
                            };

                        using Feature originFeature = (Feature)originCursor.Current;

                        // 获取源要素的几何
                        Geometry originGeometry = originFeature.GetShape();

                        // 创建空间查询过滤器，以获取与源要素有重叠的目标要素
                        SpatialQueryFilter spatialFilter = new SpatialQueryFilter
                        {
                            FilterGeometry = originGeometry,
                            SpatialRelationship = SpatialRelationship.Intersects
                        };

                        // 在目标图层中查询与源要素重叠的要素
                        using RowCursor identityCursor = originFeatureClass.Search(spatialFilter);
                        while (identityCursor.MoveNext())
                        {
                            using Feature identityFeature = (Feature)identityCursor.Current;

                            // 获取周边图斑的颜色名称
                            var color = identityFeature[field];
                            // 如果有的话，就在默认列表中移除该颜色
                            if (color is not null)
                            {
                                string colorValue = color.ToString();
                                if (colors.Contains(colorValue))
                                {
                                    colors.Remove(colorValue);
                                }
                            }

                        }

                        // 在剩余的颜色中选一个填上
                        originFeature[field] = colors[0];
                        originFeature.Store();
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

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "";
            UITool.Link2Web(url);
        }
    }
}
