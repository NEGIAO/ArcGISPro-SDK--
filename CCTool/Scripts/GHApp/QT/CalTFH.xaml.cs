using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.Util;
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
using Polygon = ArcGIS.Core.Geometry.Polygon;

namespace CCTool.Scripts.GHApp.QT
{
    /// <summary>
    /// Interaction logic for CalTFH.xaml
    /// </summary>
    public partial class CalTFH : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public CalTFH()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "计算图幅号";

        private void combox_tfhField_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_tfhField);
        }

        private void combox_blcField_DropDown(object sender, EventArgs e)
        {
            UITool.AddFieldsToComboxPlus(combox_fc.ComboxText(), combox_blcField);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayerAndTableToComboxPlus(combox_fc);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 参数获取
                string in_data = combox_fc.ComboxText();
                string tfhField = combox_tfhField.ComboxText();
                string blcField = combox_blcField.ComboxText();

                // 判断参数是否选择完全
                if (in_data == "" || tfhField == "" || blcField == "")
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
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(in_data, blcField);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10, "计算图幅号");

                    // 计算图幅号
                    using RowCursor rowCursor = in_data.TargetTable().Search();
                    while (rowCursor.MoveNext())
                    {
                        using Feature feature = rowCursor.Current as Feature;

                        string TFH = "";
                        // 比例尺
                        long blc = long.Parse(feature[blcField].ToString());

                        // 获取面
                        Polygon polygon = feature.GetShape() as Polygon;
                        // 转成地理坐标系（默认是CGCS2000）
                        SpatialReference sr = SpatialReferenceBuilder.CreateSpatialReference(4490);
                        Polygon projectPolygon = GeometryEngine.Instance.Project(polygon, sr) as Polygon;
                        // 获取范围
                        Envelope extent = projectPolygon.Extent;
                        // 获取面的四至点
                        List<List<double>> extentXY = new()
                        {
                            new List<double>(){ extent.XMin, extent.YMax },
                            new List<double>(){ extent.XMin, extent.YMin },
                            new List<double>(){ extent.XMax, extent.YMax },
                            new List<double>(){ extent.XMax, extent.YMin },
                        };
                        // 四至点都计算图幅号
                        foreach (List<double> XY in extentXY)
                        {
                            // XY坐标
                            double xx = XY[0];
                            double yy = XY[1];
                            // 计算图幅号
                            string tf = CalulateTFH(xx, yy, blc);
                            // 纳入图幅号字段
                            if (!TFH.Contains(tf))
                            {
                                TFH += $"{tf};";
                            }
                        }

                        feature[tfhField] = TFH[..^1];     // 去掉最后一个符号
                        feature.Store();
                    }


                    pw.AddMessageEnd();
                });
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/143844160";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string in_data,  string blcField)
        {
            List<string> result = new List<string>();

            // 检查字段值是否为空
            string fieldEmptyResult = CheckTool.CheckFieldValueSpace(in_data, blcField);
            if (fieldEmptyResult != "")
            {
                result.Add(fieldEmptyResult);
            }

            return result;
        }


        // 从经纬度计算图幅号
        public static string CalulateTFH(double lng, double lat, long blc)
        {
            // 1：100万行号
            int row = (int)(lat / 4 + 1);
            string rowStr = GlobalData.excelPairs[row];
            // 1：100万列号
            int col = (int)(lng / 6 + 31);
            // 比例尺代码
            Dictionary<long, string> scale = new Dictionary<long, string>()
            {
                {500000, "B"},{250000, "C"},{100000, "D"},{50000, "E"},{25000, "F"},{10000, "G"},{5000, "H"},{2000, "I"},{1000, "J"},{500, "K"},
            };
            // 纬差
            Dictionary<long, double> latDis = new Dictionary<long, double>()
            {
                {500000, 0.5},{250000,1},{100000, 0.3333},{50000, 0.1667},{25000,0.0833},{10000, 0.04167},{5000,0.020833},{2000,0.00694},{1000,0.003472},{500, 0.00173661},
            };
            // 经差
            Dictionary<long, double> lngDis = new Dictionary<long, double>()
            {
                {500000, 1},{250000,1.5},{100000, 0.5},{50000, 0.25},{25000,0.125},{10000, 0.0625},{5000,0.03125},{2000,0.010417},{1000,0.0052083},{500, 0.00260389},
            };

            // 大图幅左上坐标
            int lng_new = (col - 31) * 6;
            int lat_new = row * 4;
            // 1：5000行号
            int row_small = (int)((lat_new - lat) / latDis[blc]) + 1;
            int col_small = (int)((lng - lng_new) / lngDis[blc]) + 1;

            string row_small_str = row_small.ToString().PadLeft(3, '0');
            string col_small_str = col_small.ToString().PadLeft(3, '0');
            // 图幅号
            string tfh = $"{rowStr}{col}{scale[blc]}{row_small_str}{col_small_str}";

            return tfh;
        }
    }
}
