using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Editing.Attributes;
using ArcGIS.Desktop.Extensions;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Dialogs;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using ArcGIS.Desktop.Mapping.Events;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace CCTool.Scripts.FeaturePross
{
    /// <summary>
    /// Interaction logic for LengthStatistics.xaml
    /// </summary>
    public partial class LengthStatistics : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public LengthStatistics()
        {
            InitializeComponent();

            combox_digit.Items.Add("1");
            combox_digit.Items.Add("2");
            combox_digit.Items.Add("3");
            combox_digit.Items.Add("4");
            combox_digit.Items.Add("5");
            combox_digit.Items.Add("6");
            combox_digit.SelectedIndex = 1;

            InitiLength(2);

            // 订阅地图选择更改事件
            MapSelectionChangedEvent.Subscribe(OnMapSelectionChanged);
        }


        private void OnMapSelectionChanged(MapSelectionChangedEventArgs args)
        {
            int digit = int.Parse(combox_digit.Text);
            // 当选择更改时，自动刷新面积计算
            InitiLength(digit);
        }


        // 统计面积
        private async void InitiLength(int digit)
        {
            // 初始化变量以存储面要素的数量和各类面积指标
            int polylineCount = 0;
            double polylineLength = 0;
            double geoLength = 0;

            // 立个Flag，面要素是否有坐标系
            bool has_geo = true;
            try
            {
                await QueuedTask.Run(() =>
                {
                    // 获取活动地图视图中选定的要素集合
                    var selectedSet = MapView.Active.Map.GetSelection();

                    // 将选定的要素集合转换为字典形式
                    var selectedList = selectedSet.ToDictionary();
                    // 创建一个新的 Inspector 对象以检索要素属性
                    var inspector = new Inspector();

                    // 遍历每个选定图层及其关联的对象 ID
                    foreach (var layer in selectedList)
                    {
                        // 获取图层和关联的对象 ID
                        MapMember mapMember = layer.Key;
                        List<long> oids = layer.Value;
                        // 使用当前图层的第一个对象 ID 加载 Inspector
                        inspector.Load(mapMember, oids[0]);
                        // 获取选定要素的几何类型
                        var geometryType = inspector.Shape.GeometryType;

                        // 检查几何类型是否为面要素
                        if (geometryType == GeometryType.Polyline)
                        {
                            // 遍历当前图层中的每个对象 ID
                            foreach (var oid in oids)
                            {
                                // 使用当前对象 ID 加载 Inspector
                                inspector.Load(mapMember, oid);
                                // 将要素转换为多边形
                                var polyline = inspector.Shape as Polyline;
                                // 计算并累加多边形的面积
                                polylineLength += Math.Abs(polyline.Length);

                                if (has_geo)       // 如果坐标系还都正确的情况下
                                {
                                    //获取面要素的坐标系
                                    var sr2 = polyline.SpatialReference;
                                    if (sr2.Name == "Unknown") { has_geo = false; }        // 如果出现不正确的坐标系，后面就不要计算了
                                    else { geoLength += Math.Abs(GeometryEngine.Instance.GeodesicLength(polyline)); }       // 否则，计算椭球长度
                                }
                                // 增加面要素的数量
                                polylineCount++;
                            }
                        }
                    }
                });

                // 显示结果
                text_len_m.Text = Math.Round(polylineLength, digit).RoundWithFill(digit);
                text_len_km.Text = Math.Round(polylineLength / 1000, digit).RoundWithFill(digit);
                

                // 默认先隐藏椭球面积的信息
                lb_1.Visibility = System.Windows.Visibility.Hidden;
                lb_2.Visibility = System.Windows.Visibility.Hidden;

                text_geolen_m.Visibility = System.Windows.Visibility.Hidden;
                text_geolen_km.Visibility = System.Windows.Visibility.Hidden;


                lb_count.Content = "所选要素数量为：" + polylineCount.ToString();

                // 如果有椭球面积
                if (has_geo)
                {
                    text_geolen_m.Text = Math.Round(geoLength, digit).RoundWithFill(digit);
                    text_geolen_km.Text = Math.Round(geoLength / 1000, digit).RoundWithFill(digit);
                    // 显示椭球面积的信息
                    lb_1.Visibility = System.Windows.Visibility.Visible;
                    lb_2.Visibility = System.Windows.Visibility.Visible;

                    text_geolen_m.Visibility = System.Windows.Visibility.Visible;
                    text_geolen_km.Visibility = System.Windows.Visibility.Visible;

                    // 隐藏警告信息
                    lb_warning.Visibility = System.Windows.Visibility.Hidden;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }


        private void combox_digit_closed(object sender, EventArgs e)
        {
            int digit = int.Parse(combox_digit.Text);
            InitiLength(digit);
        }

    }
}
