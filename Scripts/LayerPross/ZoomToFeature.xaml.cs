using ActiproSoftware.Windows.Shapes;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.Geometry;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using MathNet.Numerics.LinearAlgebra.Factorization;
using NPOI.OpenXmlFormats.Vml;
using NPOI.Util;
using System;
using System.Collections;
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

namespace CCTool.Scripts.LayerPross
{
    /// <summary>
    /// Interaction logic for ZoomToFeature.xaml
    /// </summary>
    public partial class ZoomToFeature : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public ZoomToFeature()
        {
            InitializeComponent();
        }

        // 初始化OID
        long initOID = 0;

        private async void btn_next_click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 比例
                double sc = double.Parse(txt_size.Text) / 100;

                await QueuedTask.Run(() =>
                {
                    // 获取图层
                    FeatureLayer ly = MapView.Active.GetSelectedLayers().FirstOrDefault() as FeatureLayer;
                    // 获取OID字段
                    string oidField = ly.TargetIDFieldName();
                    // 获取OID值列表
                    List<string> stringList = ly.GetFieldValues(oidField);
                    List<long> oidList = stringList.Select(long.Parse).OrderBy(i => i).ToList();

                    // 如果当前没有选择，就获取【所有】要素的第一个的OID
                    if (ly.GetSelection().GetCount() == 0)
                    {
                        initOID = oidList.FirstOrDefault();
                    }
                    // 如果有选择，就获取【所选】要素的下一个的OID
                    else
                    {
                        // 获取当前选择的第一个要素的OID
                        initOID = ly.GetSelection().GetObjectIDs().ToList().FirstOrDefault();
                        // 更新initOID，获取下一个OID
                        int index = oidList.IndexOf(initOID);
                        if (index != oidList.Count - 1)
                        {
                            initOID = oidList[index + 1];
                        }

                    }

                    // 标签显示当前OID
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        txt_id.Text = @$"{initOID}/{oidList.Count}";
                    });

                    // 遍历面要素类中的所有要素
                    RowCursor cursor = ly.Search();

                    while (cursor.MoveNext())
                    {
                        using var feature = cursor.Current as Feature;
                        int oid = int.Parse(feature[oidField].ToString());
                        if (oid == initOID)
                        {
                            // 获取要素的几何
                            ArcGIS.Core.Geometry.Polygon geometry = feature.GetShape() as ArcGIS.Core.Geometry.Polygon;
                            // 选择
                            QueryFilter queryFilter = new QueryFilter();
                            queryFilter.WhereClause = $"{oidField} = {oid}";
                            ly.Select(queryFilter);
                            // 缩放至图斑
                            MapCtlTool.Zoom2Feature(feature, sc);

                            break;
                        }

                    }
                });
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        private async void btn_last_click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 比例
                double sc = double.Parse(txt_size.Text)/100;

                await QueuedTask.Run(() =>
                {
                    // 获取图层
                    FeatureLayer ly = MapView.Active.GetSelectedLayers().FirstOrDefault() as FeatureLayer;
                    // 获取OID字段
                    string oidField = ly.TargetIDFieldName();
                    // 获取OID值列表
                    List<string> stringList = ly.GetFieldValues(oidField);
                    List<long> oidList = stringList.Select(long.Parse).OrderBy(i => i).ToList();

                    // 如果当前没有选择，就获取【所有】要素的最后一个的OID
                    if (ly.GetSelection().GetCount() == 0)
                    {
                        initOID = oidList.LastOrDefault();
                    }
                    // 如果有选择，就获取【所选】要素的上一个的OID
                    else
                    {
                        // 获取当前选择的第一个要素的OID
                        initOID = ly.GetSelection().GetObjectIDs().ToList().FirstOrDefault();
                        // 更新initOID，获取上一个OID
                        int index = oidList.IndexOf(initOID);
                        if (index != 0)
                        {
                            initOID = oidList[index - 1];
                        }

                    }

                    // 标签显示当前OID
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        txt_id.Text = @$"{initOID}/{oidList.Count}";
                    });

                    // 遍历面要素类中的所有要素
                    RowCursor cursor = ly.Search();

                    while (cursor.MoveNext())
                    {
                        using var feature = cursor.Current as Feature;
                        int oid = int.Parse(feature[oidField].ToString());
                        if (oid == initOID)
                        {
                            // 选择
                            QueryFilter queryFilter = new QueryFilter();
                            queryFilter.WhereClause = $"{oidField} = {oid}";
                            ly.Select(queryFilter);
                            // 缩放至图斑
                            MapCtlTool.Zoom2Feature(feature, sc);

                            break;
                        }

                    }
                });
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }
    }
}
