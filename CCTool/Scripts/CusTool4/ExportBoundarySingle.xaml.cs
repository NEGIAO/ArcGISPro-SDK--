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
using System.Security.Cryptography;
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
using Polygon = ArcGIS.Core.Geometry.Polygon;
using Range = Aspose.Cells.Range;

namespace CCTool.Scripts.CusTool4
{
    /// <summary>
    /// Interaction logic for ExportBoundarySingle.xaml
    /// </summary>
    public partial class ExportBoundarySingle : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "ExportBoundarySingle";
        public ExportBoundarySingle()
        {
            InitializeComponent();

            combox_digit.Items.Add("1");
            combox_digit.Items.Add("2");
            combox_digit.Items.Add("3");
            combox_digit.Items.Add("4");
            combox_digit.Items.Add("5");
            combox_digit.Items.Add("6");
            combox_digit.SelectedIndex = 3;

            // 初始化参数选项
            textExcelPath.Text = BaseTool.ReadValueFromReg(toolSet, "excelPath");

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "界址点导出单个Excel(牛)";


        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.SaveDialogXls();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 参数获取
                string in_fc = combox_fc.ComboxText();
                string excelPath = textExcelPath.Text;

                int ptDigit = combox_digit.Text.ToInt();

                // 判断参数是否选择完全
                if (in_fc == "" || excelPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 判断参数是否选择完全
                if (excelPath.Contains(".xlsx"))
                {
                    MessageBox.Show("保存文件的后缀格式须为.xls");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "excelPath", excelPath);


                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    // 复制界址点Excel表
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.杂七杂八.压矿Excel导入模板.xls", excelPath);
                    // 获取工作薄、工作表
                    string excelFile = ExcelTool.GetPath(excelPath);
                    int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
                    // 打开工作薄
                    Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                    // 打开工作表
                    Worksheet worksheet = wb.Worksheets[sheetIndex];
                    // 获取Cells
                    Cells cells = worksheet.Cells;

                    pw.AddMessageStart("获取目标FeatureLayer");
                    // 获取目标FeatureLayer
                    FeatureLayer featurelayer = in_fc.TargetFeatureLayer();
                    // 遍历面要素类中的所有要素
                    RowCursor cursor = featurelayer.TargetSelectCursor();
                    long featureCount = featurelayer.GetFeatureCount();
                    int dkh = 1;  // 地块圈号
                    int rowIndex = 1;   // 起始行
                    pw.AddMessageMiddle(20, $"总图斑数：{featureCount}");

                    while (cursor.MoveNext())
                    {
                        using var feature = cursor.Current as Feature;
                        // 获取要素的几何
                        Polygon geometry = feature.GetShape() as Polygon;
                        if (geometry != null)
                        {
                            // 获取面要素的所有折点【按西北角起始，顺时针重排】
                            List<List<MapPoint>> mapPoints = geometry.ReshotMapPoint();
                            int pointIndex = 1;  // 起始点序号
                            int lastRowCount = 0;  // 上一轮的行数

                            for (int i = 0; i < mapPoints.Count; i++)
                            {
                                for (int j = 0; j < mapPoints[i].Count; j++)
                                {
                                    if (rowIndex>1)
                                    {
                                        cells.CopyRow(cells, 1, rowIndex);
                                    }

                                    // 写入点号
                                    if (pointIndex - lastRowCount == mapPoints[i].Count)    // 找到当前环的最后一点
                                    {
                                        worksheet.Cells[rowIndex, 0].Value = $"J{lastRowCount + 1 - i}";
                                    }
                                    else
                                    {
                                        worksheet.Cells[rowIndex, 0].Value = $"J{pointIndex - i}";
                                    }

                                    // 写入地块号
                                    worksheet.Cells[rowIndex, 4].Value = dkh;

                                    double x = Math.Round(mapPoints[i][j].X, ptDigit);
                                    double y = Math.Round(mapPoints[i][j].Y, ptDigit);
                                    // 写入折点的XY值
                                    worksheet.Cells[rowIndex, 2].Value = x;
                                    worksheet.Cells[rowIndex, 1].Value = y;
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
                                    worksheet.Cells[rowIndex, 1].SetStyle(style);

                                    pointIndex++;
                                    rowIndex++;
                                }
                                lastRowCount += mapPoints[i].Count;
                            }
                        }

                        dkh++;
                    }

                    // 保存
                    wb.Save(excelPath);
                    wb.Dispose();

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
            string url = "https://blog.csdn.net/xcc34452366/article/details/148089079";
            UITool.Link2Web(url);
        }
    }
}
