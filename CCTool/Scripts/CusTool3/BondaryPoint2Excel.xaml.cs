using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Framework.Utilities;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using MathNet.Numerics;
using MathNet.Numerics.LinearAlgebra.Factorization;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
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
using System.Windows.Shapes;
using Range = Aspose.Cells.Range;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for BondaryPoint2Excel.xaml
    /// </summary>
    public partial class BondaryPoint2Excel : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "BondaryPoint2Excel";

        public BondaryPoint2Excel()
        {
            InitializeComponent();
            // 初始化combox
            combox_zd_unit.Items.Add("平方米");
            combox_zd_unit.Items.Add("公顷");
            combox_zd_unit.Items.Add("平方公里");
            combox_zd_unit.Items.Add("亩");
            combox_zd_unit.SelectedIndex = 0;

            combox_zd_areaDigit.Items.Add("0");
            combox_zd_areaDigit.Items.Add("1");
            combox_zd_areaDigit.Items.Add("2");
            combox_zd_areaDigit.Items.Add("3");
            combox_zd_areaDigit.Items.Add("4");
            combox_zd_areaDigit.Items.Add("5");
            combox_zd_areaDigit.Items.Add("6");
            combox_zd_areaDigit.SelectedIndex = 2;

            combox_ptDigit.Items.Add("1");
            combox_ptDigit.Items.Add("2");
            combox_ptDigit.Items.Add("3");
            combox_ptDigit.Items.Add("4");
            combox_ptDigit.SelectedIndex = 2;

            combox_lenDigit.Items.Add("1");
            combox_lenDigit.Items.Add("2");
            combox_lenDigit.Items.Add("3");
            combox_lenDigit.Items.Add("4");
            combox_lenDigit.SelectedIndex = 1;

            combox_zd_type.Items.Add("椭球面积");
            combox_zd_type.Items.Add("投影面积");
            combox_zd_type.SelectedIndex = 1;

            combox_sh.Items.Add("按点号排序");
            combox_sh.Items.Add("按部件号排序");
            combox_sh.SelectedIndex = 0;

            // 初始化参数选项
            textFolderPath.Text = BaseTool.ReadValueFromReg(toolSet, "excel_folder");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "界址点导出Excel";

        private void combox_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textFolderPath.Text = UITool.OpenDialogFolder();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 参数获取
                string in_fc = combox_fc.ComboxText();
                string zdh_field = combox_field.ComboxText();
                string excel_folder = textFolderPath.Text;
                string zd_unit = combox_zd_unit.Text;
                int zd_areaDigit = int.Parse(combox_zd_areaDigit.Text);
                int ptDigit = int.Parse(combox_ptDigit.Text);
                int lenDigit = int.Parse(combox_lenDigit.Text);

                string qlr = combox_qlr.ComboxText();
                bool xyReserve = (bool)check_xy.IsChecked;
                bool haveJ = (bool)check_xy_J.IsChecked;
                string zdmj_type = combox_zd_type.Text;
                string sh = combox_sh.Text;     //序号模式

                // 判断参数是否选择完全
                if (in_fc == "" || zdh_field == "" || excel_folder == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "excel_folder", excel_folder);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("获取目标FeatureLayer");
                    // 获取目标FeatureLayer
                    FeatureLayer featurelayer = in_fc.TargetFeatureLayer();

                    // 遍历面要素类中的所有要素
                    RowCursor cursor = featurelayer.TargetSelectCursor();
                    while (cursor.MoveNext())
                    {
                        using var feature = cursor.Current as Feature;
                        // 获取要素的几何
                        ArcGIS.Core.Geometry.Polygon geometry = feature.GetShape() as ArcGIS.Core.Geometry.Polygon;
                        // 获取ID\名称\权利人
                        string oidField = in_fc.TargetIDFieldName();
                        string oid = feature[oidField].ToString();
                        string feature_name = feature[zdh_field]?.ToString();   // 宗地号

                        pw.AddMessageMiddle(20, $"处理要素：{oid} - {feature_name}");

                        string QLR = "";                    // 权利人
                        if (qlr != "") { QLR = feature[qlr]?.ToString() ?? ""; }

                        // 计算多边形的面积
                        double xs = zd_unit switch    // 单位换算
                        {
                            "平方米" => 1,
                            "公顷" => 10000,
                            "平方公里" => 1000000,
                            "亩" => 666.6667,
                            _ => 0,
                        };
                        double polygonArea;
                        if (zdmj_type == "椭球面积")
                        {
                            polygonArea = Math.Round(GeometryEngine.Instance.GeodesicArea(geometry) / xs, zd_areaDigit);
                        }
                        else
                        {
                            polygonArea = Math.Round(geometry.Area / xs, zd_areaDigit);
                        }

                        // 复制界址点Excel表
                        string excelPath = excel_folder + @$"\{oid} - {feature_name}界址点表.xlsx";
                        DirTool.CopyResourceFile(@"CCTool.Data.Excel.杂七杂八.3 - 运输便道100界址点表.xlsx", excelPath);
                        // 获取工作薄、工作表
                        string excelFile = ExcelTool.GetPath(excelPath);
                        int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
                        // 打开工作薄
                        Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                        // 打开工作表
                        Worksheet worksheet = wb.Worksheets[sheetIndex];
                        // 获取Cells
                        Cells cells = worksheet.Cells;


                        // 设置分页信息
                        SetPage(geometry, cells, feature_name, QLR, polygonArea, zd_unit);

                        if (geometry != null)
                        {
                            // 获取面要素的所有折点【按西北角起始，顺时针重排】
                            List<List<MapPoint>> mapPoints = geometry.ReshotMapPoint();

                            int rowIndex = 7;   // 起始行
                            int pointIndex = 1;  // 起始点序号
                            int lastRowCount = 0;  // 上一轮的行数

                            for (int i = 0; i < mapPoints.Count; i++)
                            {
                                // 输出折点的XY值和距离到Excel表
                                double prevX = double.NaN;
                                double prevY = double.NaN;

                                for (int j = 0; j < mapPoints[i].Count; j++)
                                {
                                    // 写入点号
                                    // J前缀
                                    string jFront = haveJ switch
                                    {
                                        true => "J",
                                        false => "",
                                    };
                                    if (pointIndex - lastRowCount == mapPoints[i].Count)    // 找到当前环的最后一点
                                    {
                                        worksheet.Cells[rowIndex, 1].Value = $"{jFront}{lastRowCount + 1 - i}";
                                    }
                                    else
                                    {
                                        worksheet.Cells[rowIndex, 1].Value = $"{jFront}{pointIndex - i}";
                                    }

                                    // 写入序号
                                    if (sh == "按部件号排序")
                                    {
                                        worksheet.Cells[rowIndex, 0].Value = $"{i + 1}";
                                    }
                                    else if (sh == "按点号排序")
                                    {
                                        worksheet.Cells[rowIndex, 0].Value = $"{pointIndex}";
                                    }

                                    double x = Math.Round(mapPoints[i][j].X, ptDigit);
                                    double y = Math.Round(mapPoints[i][j].Y, ptDigit);
                                    // 写入折点的XY值
                                    if (xyReserve)
                                    {
                                        worksheet.Cells[rowIndex, 2].Value = x;
                                        worksheet.Cells[rowIndex, 3].Value = y;
                                    }
                                    else
                                    {
                                        worksheet.Cells[rowIndex, 2].Value = y;
                                        worksheet.Cells[rowIndex, 3].Value = x;
                                    }
                                    // 设置单元格为数字型，小数位数
                                    Aspose.Cells.Style style = worksheet.Cells[rowIndex, 2].GetStyle();
                                    style.Number = 4;   // 数字型
                                    // 小数位数
                                    style.Custom = ptDigit switch
                                    {
                                        1 => "0.0",
                                        2 => "0.00",
                                        3 => "0.000",
                                        4 => "0.0000",
                                        _ => null,
                                    };
                                    // 设置
                                    worksheet.Cells[rowIndex, 2].SetStyle(style);
                                    worksheet.Cells[rowIndex, 3].SetStyle(style);

                                    // 计算当前点与上一个点的距离
                                    if (!double.IsNaN(prevX) && !double.IsNaN(prevY) && rowIndex > 8)
                                    {
                                        string distance = Math.Sqrt(Math.Pow(x - prevX, 2) + Math.Pow(y - prevY, 2)).RoundWithFill(lenDigit);
                                        // 写入距离【第一行不写入】
                                        worksheet.Cells[rowIndex - 1, 4].Value = distance;
                                        if (j == mapPoints[i].Count - 1)
                                        {
                                            worksheet.Cells[rowIndex + 1, 4].Value = "";
                                        }
                                    }
                                    prevX = x;
                                    prevY = y;
                                    // 是否跨页
                                    if (rowIndex == 79 || (rowIndex - 79) % 85 == 0)
                                    {
                                        // 但是如果不是最后一个点，就回退
                                        if (pointIndex - lastRowCount != mapPoints[i].Count)
                                        {
                                            rowIndex += 11;
                                            // 点号回退1
                                            j--;
                                            pointIndex--;
                                        }
                                        else
                                        {
                                            rowIndex += 11;
                                        }
                                    }
                                    else
                                    {
                                        rowIndex += 2;
                                    }
                                    pointIndex++;
                                }
                                lastRowCount += mapPoints[i].Count;
                            }
                        }
                        // 保存
                        wb.Save(excelPath);
                        wb.Dispose();
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

        // 生成分页，并设置分页信息
        public static void SetPage(ArcGIS.Core.Geometry.Polygon geometry, Cells cells, string ZDH, string QLR, double ZDMJ, string zd_unit)
        {
            // 获取要素的界址点总数
            var pointsCount = geometry.Points.Count;

            // 计算要生成的页数
            long pageCount = (int)Math.Ceiling((double)pointsCount / 37);

            if (pageCount == 1)   // 只有一页，就删第二页
            {
                cells.DeleteRows(83, 85);
            }
            else if (pageCount > 2)  // 多于2页
            {
                // 复制分页，前两页已有不用复制
                for (int i = 0; i < pageCount - 2; i++)
                {
                    cells.CopyRows(cells, 83, 168 + i * 85, 85);
                }
            }
            // 设置单元格
            for (int i = 0; i < pageCount; i++)
            {
                if (i == 0) // 第一页
                {
                    // 填写页码
                    cells[0, 4].Value = "第 1 页";
                    cells[1, 4].Value = $"共 {pageCount} 页";
                    // 宗地号
                    string zdhStr = cells[0, 0].StringValue;
                    cells[0, 0].Value = zdhStr.Replace("运输便道100", ZDH);
                    // 权利人
                    string qlrStr = cells[2, 0].StringValue;
                    cells[2, 0].Value = qlrStr.Replace("QLR", QLR);
                    // 宗地面积
                    string zdmjStr = cells[3, 0].StringValue;
                    cells[3, 0].Value = zdmjStr.Replace("ZDMJ", ZDMJ.ToString()).Replace("平方米", zd_unit);

                }
                else  // 其它页
                {
                    // 填写页码
                    cells[83 + (i - 1) * 85, 4].Value = $"第 {i + 1} 页";
                    cells[84 + (i - 1) * 85, 4].Value = $"共 {pageCount} 页";
                    // 宗地号
                    string zdhStr = cells[83 + (i - 1) * 85, 0].StringValue;
                    cells[83 + (i - 1) * 85, 0].Value = zdhStr.Replace("运输便道100", ZDH);
                    // 权利人
                    string qlrStr = cells[85 + (i - 1) * 85, 0].StringValue;
                    cells[85 + (i - 1) * 85, 0].Value = qlrStr.Replace("QLR", QLR);
                    // 宗地面积
                    string zdmjStr = cells[86 + (i - 1) * 85, 0].StringValue;
                    cells[86 + (i - 1) * 85, 0].Value = zdmjStr.Replace("ZDMJ", ZDMJ.ToString()).Replace("平方米", zd_unit);

                }

            }
        }

        private void combox_qlr_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_qlr);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135836514?spm=1001.2014.3001.5502";
            UITool.Link2Web(url);
        }
    }
}
