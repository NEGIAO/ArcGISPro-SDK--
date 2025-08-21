using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Utilities;
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

namespace CCTool.Scripts.CusTool
{
    /// <summary>
    /// Interaction logic for ExportBound2Excel.xaml
    /// </summary>
    public partial class ExportBound2Excel : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public ExportBound2Excel()
        {
            InitializeComponent();

            // 初始化combox
            combox_ptDigit.Items.Add("1");
            combox_ptDigit.Items.Add("2");
            combox_ptDigit.Items.Add("3");
            combox_ptDigit.Items.Add("4");
            combox_ptDigit.SelectedIndex = 2;

            UITool.InitFieldToComboxPlus(combox_mj, "Shape_Area", "float");

            string bz = BaseTool.ReadValueFromReg("BoundExcel", "编制");
            string fh = BaseTool.ReadValueFromReg("BoundExcel", "复核");

            txt_zb.Text = bz;
            txt_js.Text = fh;
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "界址点导出Excel表(名特长)";

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
                string excel_folder = textFolderPath.Text;

                string xmmcField = combox_xmmc.ComboxText();
                string mjField = combox_mj.ComboxText();

                string zb = txt_zb.Text;
                string js = txt_js.Text;

                int ptDigit = int.Parse(combox_ptDigit.Text);
                bool xyReserve = (bool)check_xy.IsChecked;
                bool haveJ = (bool)check_xy_J.IsChecked;

                // 保留预置信息
                BaseTool.WriteValueToReg("BoundExcel", "编制", zb);
                BaseTool.WriteValueToReg("BoundExcel", "复核", js);

                // 判断参数是否选择完全
                if (in_fc == "" || excel_folder == "")
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
                    pw.AddMessageStart("获取目标FeatureLayer");
                    // 获取目标FeatureLayer
                    FeatureLayer featurelayer = in_fc.TargetFeatureLayer();
                    // 确保要素类的几何类型是多边形
                    if (featurelayer.ShapeType != esriGeometryType.esriGeometryPolygon)
                    {
                        // 如果不是多边形类型，则输出错误信息并退出函数
                        MessageBox.Show("该要素类不是多边形类型。");
                        return;
                    }
                    string oidField = in_fc.TargetIDFieldName();

                    // 遍历面要素类中的所有要素
                    RowCursor cursor = featurelayer.TargetSelectCursor();

                    while (cursor.MoveNext())
                    {
                        using var feature = cursor.Current as Feature;
                        // 获取要素的几何
                        ArcGIS.Core.Geometry.Polygon geometry = feature.GetShape() as ArcGIS.Core.Geometry.Polygon;
                        // 获取参数
                        string xmmc = feature[xmmcField]?.ToString();   // 项目名称
                        double mj = double.Parse(feature[mjField]?.ToString());   // 宗地面积
                        double mj_m = mj / 10000 * 15;   // 宗地面积_亩
                        // 转小数点后2位
                        string zdmj = mj.RoundWithFill(2);
                        string zdmj_m = mj_m.RoundWithFill(2);

                        var oid = feature[oidField]?.ToString();   // OID

                        pw.AddMessageMiddle(20, $"处理要素: {oid}-{xmmc}");
                        // 复制界址点Excel表
                        string excelPath = excel_folder + @$"\{oid}-{xmmc}-界址点表.xlsx";
                        DirTool.CopyResourceFile(@"CCTool.Data.Excel.【模板】界址点表_多页_自定义_名特长.xlsx", excelPath);
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
                        SetPage(geometry, cells, xmmc, zdmj, zdmj_m, zb, js);

                        if (geometry != null)
                        {
                            // 获取面要素的所有折点【按西北角起始，顺时针重排】
                            List<List<MapPoint>> mapPoints = geometry.ReshotMapPoint();

                            int rowIndex = 8;   // 起始行
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
                                    worksheet.Cells[rowIndex, 0].Value = $"{pointIndex}";


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
                                    Aspose.Cells.Style style = worksheet.Cells[rowIndex, 3].GetStyle();
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
                                        double distance = Math.Round(Math.Sqrt(Math.Pow(x - prevX, 2) + Math.Pow(y - prevY, 2)), 2);
                                        // 写入距离【第一行不写入】
                                        worksheet.Cells[rowIndex - 1, 4].Value = distance;
                                        if (j == mapPoints[i].Count - 1)
                                        {
                                            worksheet.Cells[rowIndex + 1, 4].Value = "";
                                        }
                                    }
                                    prevX = x;
                                    prevY = y;

                                    // 设置点距格式
                                    Aspose.Cells.Style style2 = worksheet.Cells[rowIndex, 4].GetStyle();
                                    style2.Number = 4;   // 数字型
                                    // 小数位数
                                    style2.Custom = 2 switch
                                    {
                                        1 => "0.0",
                                        2 => "0.00",
                                        3 => "0.000",
                                        4 => "0.0000",
                                        _ => null,
                                    };
                                    // 设置
                                    worksheet.Cells[rowIndex + 1, 4].SetStyle(style2);

                                    // 是否跨页
                                    if ((rowIndex - 76) % 80 == 0)
                                    {
                                        rowIndex += 12;
                                        // 如果不是最后一个点，点号回退
                                        if (pointIndex - lastRowCount != mapPoints[i].Count)
                                        {
                                            // 点号回退1
                                            j--;
                                            pointIndex--;
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
        public static void SetPage(ArcGIS.Core.Geometry.Polygon geometry, Cells cells, string xmmc, string zdmj, string zdmj_m, string zb, string js)
        {
            // 获取要素的界址点总数
            var pointsCount = geometry.Points.Count;
            // 计算要生成的页数
            long pageCount = (int)Math.Ceiling((double)(pointsCount - 35) / 34) + 1;

            int rowsInPage = 80;    // 每页的行数

            if (pageCount > 1)  // 多于1页
            {
                // 复制分页
                for (int i = 1; i < pageCount; i++)
                {
                    cells.CopyRows(cells, 0, i * rowsInPage, rowsInPage);
                }
            }
            // 设置单元格
            for (int i = 0; i < pageCount; i++)
            {
                // 填写页码
                cells[i * rowsInPage, 4].Value = $"第 {i + 1} 页";
                cells[i * rowsInPage + 1, 4].Value = $"共 {pageCount} 页";
                // 项目名称、地块面积
                cells[i * rowsInPage + 2, 0].Value = $"{xmmc} 面积:{zdmj}平方米，合{zdmj_m}亩\n换算系数：（666.67）";

                // 编制
                cells[rowsInPage * (i + 1) - 2, 0].Value = $"编制：{zb}";
                // 复核
                cells[rowsInPage * (i + 1) - 2, 2].Value = $"复核：{js}";
            }
        }


        private void combox_xmmc_DropOpen(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_xmmc);
        }

        private void combox_mj_DropOpen(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_mj);
        }
    }
}
