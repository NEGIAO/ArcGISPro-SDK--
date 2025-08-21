using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using System;
using System.Collections.Generic;
using System.IO;
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
using Geometry = ArcGIS.Core.Geometry.Geometry;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for PickJDFeature.xaml
    /// </summary>
    public partial class PickJDFeature : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "PickJDFeature";
        public PickJDFeature()
        {
            InitializeComponent();

            //UITool.InitFeatureLayerToComboxPlus(combox_jd, "宗地");
            //UITool.InitFeatureLayerToComboxPlus(combox_jzd, "JZD");
            //UITool.InitFeatureLayerToComboxPlus(combox_dltb, "DLTB");

            // 初始化参数选项
            textFolderPath.Text = BaseTool.ReadValueFromReg(toolSet, "folderPath");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "整理宗地要素(伊)";

        private void combox_jd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_jd);
        }

        private void combox_jzd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_jzd, "Point");
        }

        private void combox_dltb_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_dltb);
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textFolderPath.Text = UITool.OpenDialogFolder();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                string defGDB = Project.Current.DefaultGeodatabasePath;
                string defFolder = Project.Current.HomeFolderPath;

                // 获取参数
                string jd = combox_jd.ComboxText();
                string jzd = combox_jzd.ComboxText();
                string dltb = combox_dltb.ComboxText();
                string folderPath = textFolderPath.Text;


                // 判断参数是否选择完全
                if (jd == "" || jzd == "" || dltb == "" || folderPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "folderPath", folderPath);

                // 创建目标文件夹
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(jd, jzd);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }
                    // 获取图层
                    FeatureLayer featureLayer = jd.TargetFeatureLayer();
                    // 要素个数
                    long count = featureLayer.GetFeatureCount();

                    pw.AddMessageMiddle(0, "按单元号导出");

                    using RowCursor rowCursor = featureLayer.Search();
                    while (rowCursor.MoveNext())
                    {
                        Feature feature = rowCursor.Current as Feature;
                        Polygon polygon = feature.GetShape() as Polygon;

                        // 获取字段值
                        string dyh = feature["BDCDYH"]?.ToString();
                        string zddm = feature["ZDDM"]?.ToString();

                        pw.AddMessageMiddle(80 / count, $"    {dyh}", Brushes.Gray);

                        // 创建宗地文件夹
                        string dyFolder = $@"{folderPath}\{dyh}";
                        DirTool.CreateFolder(dyFolder);

                        // 选择当前宗地图斑
                        QueryFilter queryFilter = new QueryFilter();
                        queryFilter.WhereClause = $"ZDDM = '{zddm}'";
                        featureLayer.Select(queryFilter);

                        // 提取界址点
                        Arcpy.SelectLayerByLocation(jzd, featureLayer);
                        // 复制到指定位置
                        string jzdPath = $@"{dyFolder}\{zddm}JZD.shp";
                        Arcpy.CopyFeatures(jzd, jzdPath);
                        // 更新ZDDM
                        Arcpy.CalculateField(jzdPath, "ZDDM", $"'{zddm}'");
                        // 删除可能存在的重叠界址点
                        Arcpy.DeleteIdentical(jzdPath);

                        string lineZD = $@"{defGDB}\lineZD";
                        string linePath = $@"{dyFolder}\{zddm}JZX.shp";

                        // 设置起始点到任一界址点上
                        SetStartPoint(feature, polygon, jzdPath);

                        // 宗地转界址线
                        Arcpy.FeatureToLine(featureLayer, lineZD);

                        // 在界址点处打断
                        Arcpy.SplitLineAtPoint(lineZD, jzd, linePath);

                        // 三调提取
                        string dltbPath = $@"{dyFolder}\{zddm}DLTB.shp";
                        Arcpy.Clip(dltb, jd, dltbPath);
                    }

                    Arcpy.Delect($@"{defGDB}\lineZD");

                    // 恢复取消选择
                    MapCtlTool.UnSelectAllFeature(jd);
                    MapCtlTool.UnSelectAllFeature(jzd);
                });

                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private List<string> CheckData(string jd, string jzd)
        {
            List<string> result = new List<string>();

            if (jd != "")
            {
                // 检查字段值是否符合要求
                string result_value = CheckTool.CheckFieldValueSpace(jd, new List<string>() { "ZDDM", "BDCDYH" });
                if (result_value != "")
                {
                    result.Add(result_value);
                }
            }

            if (jzd != "")
            {
                // 检查字段值是否符合要求
                string result_value = CheckTool.IsHaveFieldInLayer(jzd, "ZDDM");
                if (result_value != "")
                {
                    result.Add(result_value);
                }
            }
            return result;
        }


        // 设置面的起始点
        private void SetStartPoint(Feature feature, Polygon polygon, string pointPath)
        {
            if (polygon != null)
            {
                // 找点
                List<double> xy = new List<double>();

                FeatureClass featureClass = pointPath.TargetFeatureClass();
                using RowCursor rowCursor = featureClass.Search();
                while (rowCursor.MoveNext())
                {
                    Feature ft = rowCursor.Current as Feature;
                    MapPoint point = ft.GetShape() as MapPoint;

                    xy.Add(point.X);
                    xy.Add(point.Y);

                    break;
                }

                // 面要素的所有折点进行重排【按西北角起始，顺时针重排】
                Polygon resultPolygon = polygon.ReshotMapPointReturnPolygonByCustom(xy);
                // 重新设置要素并保存
                feature.SetShape(resultPolygon);
                feature.Store();
            }
        }
    }
}
