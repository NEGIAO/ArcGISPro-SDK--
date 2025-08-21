using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Framework.Utilities;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using MathNet.Numerics;
using MathNet.Numerics.LinearAlgebra.Factorization;
using NPOI.POIFS.Crypt.Dsig;
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
    /// Interaction logic for ExportBoundaryPoints4.xaml
    /// </summary>
    public partial class ExportBoundaryPoints4 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public ExportBoundaryPoints4()
        {
            InitializeComponent();

            // 读取参数
            textFolderPath.Text = BaseTool.ReadValueFromReg("ExportBound0", "excelFolder");
            txt_xmmc.Text = BaseTool.ReadValueFromReg("ExportBound0", "xmmc");
            txt_dkwz.Text = BaseTool.ReadValueFromReg("ExportBound0", "dkwz");
            txt_jsz.Text = BaseTool.ReadValueFromReg("ExportBound0", "jsz");
            txt_jcz.Text = BaseTool.ReadValueFromReg("ExportBound0", "jcz");
            txt_zbdw.Text = BaseTool.ReadValueFromReg("ExportBound0", "zbdw");
            txt_shdw.Text = BaseTool.ReadValueFromReg("ExportBound0", "shdw");
            txt_zbrq.Text = BaseTool.ReadValueFromReg("ExportBound0", "zbrq");
            txt_sr.Text = BaseTool.ReadValueFromReg("ExportBound0", "sr");


            combox_ptDigit.Items.Add("1");
            combox_ptDigit.Items.Add("2");
            combox_ptDigit.Items.Add("3");
            combox_ptDigit.Items.Add("4");
            combox_ptDigit.SelectedIndex = 2;
        }


        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "界址点导出Excel";


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

                string dkmc_field = combox_dkmc.ComboxText();

                string xmmc_txt = txt_xmmc.Text;
                string dkwz_txt = txt_dkwz.Text;
                string jsz_txt = txt_jsz.Text;
                string jcz_txt = txt_jcz.Text;
                string zbdw_txt = txt_zbdw.Text;
                string shdw_txt = txt_shdw.Text;
                string zbrq_txt = txt_zbrq.Text;
                string sr = txt_sr.Text;

                string xmmc_field = combox_xmmc.ComboxText();
                string dkwz_field = combox_dkwz.ComboxText();
                string jsz_field = combox_jsz.ComboxText();
                string jcz_field = combox_jcz.ComboxText();
                string zbdw_field = combox_zbdw.ComboxText();
                string shdw_field = combox_shdw.ComboxText();
                string zbrq_field = combox_zbrq.ComboxText();

                bool bool_xmmc = (bool)cb_xmmc.IsChecked;
                bool bool_dkwz = (bool)cb_dkwz.IsChecked;
                bool bool_jsz = (bool)cb_jsz.IsChecked;
                bool bool_jcz = (bool)cb_jcz.IsChecked;
                bool bool_zbdw = (bool)cb_zbdw.IsChecked;
                bool bool_shdw = (bool)cb_shdw.IsChecked;
                bool bool_zbrq = (bool)cb_zbrq.IsChecked;

                int digit = combox_ptDigit.SelectedIndex + 1;

                // 记录参数
                BaseTool.WriteValueToReg("ExportBound0", "excelFolder", excel_folder);
                BaseTool.WriteValueToReg("ExportBound0", "xmmc", xmmc_txt);
                BaseTool.WriteValueToReg("ExportBound0", "dkwz", dkwz_txt);
                BaseTool.WriteValueToReg("ExportBound0", "jsz", jsz_txt);
                BaseTool.WriteValueToReg("ExportBound0", "jcz", jcz_txt);
                BaseTool.WriteValueToReg("ExportBound0", "zbdw", zbdw_txt);
                BaseTool.WriteValueToReg("ExportBound0", "shdw", shdw_txt);
                BaseTool.WriteValueToReg("ExportBound0", "zbrq", zbrq_txt);
                BaseTool.WriteValueToReg("ExportBound0", "sr", sr);

                // 判断参数是否选择完全
                if (in_fc == "" || dkmc_field == "" || excel_folder == "")
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

                    // 遍历面要素类中的所有要素
                    RowCursor cursor = featurelayer.TargetSelectCursor();
                    while (cursor.MoveNext())
                    {
                        using var feature = cursor.Current as Feature;
                        // 获取要素的几何
                        ArcGIS.Core.Geometry.Polygon polygon = feature.GetShape() as ArcGIS.Core.Geometry.Polygon;
                        // 获取ID\名称\权利人
                        string oidField = in_fc.TargetIDFieldName();
                        string oid = feature[oidField].ToString();
                        string dkmc = feature[dkmc_field]?.ToString();   // 地块名称

                        string xmmc = bool_xmmc ? xmmc_txt : feature[xmmc_field]?.ToString();
                        string dkwz = bool_dkwz ? dkwz_txt : feature[dkwz_field]?.ToString();
                        string jsz = bool_jsz ? jsz_txt : feature[jsz_field]?.ToString();

                        string jcz = bool_jcz ? jcz_txt : feature[jcz_field]?.ToString();
                        string zbdw = bool_zbdw ? zbdw_txt : feature[zbdw_field]?.ToString();

                        string shdw = bool_shdw ? shdw_txt : feature[shdw_field]?.ToString();
                        string zbrq = bool_zbrq ? zbrq_txt : feature[zbrq_field]?.ToString();


                        pw.AddMessageMiddle(20, $"处理要素：{oid} - {dkmc}");

                        // 地块名称转化为点号
                        string DH = "";

                        try
                        {
                            string numStr = dkmc.Replace("地块", "").Replace(" ", "");
                            long numInt = TxtTool.ParseCnToInt(numStr);
                            DH = TxtTool.NumberChange(numInt);

                        }
                        catch (Exception)
                        {
                            MessageBox.Show($"ID【{oid}】：地块名不规范。应如格式【地块二十一】");
                            return;
                        }


                        // 面积_平方米，公顷
                        string area_m = polygon.Area.RoundWithFill(2) + "平方米";
                        string area_ha = (polygon.Area / 10000).RoundWithFill(4) + "公顷";

                        // 界线总长
                        string lineLenth = GisTool.GetLineLength(polygon).RoundWithFill(2) + "米";
                        string pointCount = (polygon.PointCount - polygon.PartCount).ToString();

                        // 复制界址点Excel表
                        string excelPath = excel_folder + @$"\{oid}_{dkmc}_界址点表.xlsx";
                        DirTool.CopyResourceFile(@"CCTool.Data.Excel.【模板】界址点表_多页_0.xlsx", excelPath);
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
                        SetPage(polygon, cells, xmmc, dkmc, dkwz, area_m, area_ha, lineLenth, pointCount, jsz, jcz, zbdw, shdw, zbrq, sr);

                        if (polygon != null)
                        {
                            // 获取面要素的所有折点【按西北角起始，顺时针重排】
                            List<List<MapPoint>> mapPoints = polygon.ReshotMapPoint();

                            int rowIndex = 3;   // 起始行
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
                                    if (pointIndex - lastRowCount == mapPoints[i].Count)    // 找到当前环的最后一点
                                    {
                                        cells[rowIndex, 0].Value = $"J{DH}{lastRowCount + 1 - i}";
                                    }
                                    else
                                    {
                                        cells[rowIndex, 0].Value = $"J{DH}{pointIndex - i}";
                                    }

                                    double x = Math.Round(mapPoints[i][j].X, digit);
                                    double y = Math.Round(mapPoints[i][j].Y, digit);

                                    // 写入折点的XY值
                                    cells[rowIndex, 2].Value = y;
                                    cells[rowIndex, 3].Value = x;

                                    // 写入标志
                                    cells[rowIndex, 4].Value = "木桩";

                                    // 设置单元格为数字型，小数位数
                                    Aspose.Cells.Style style = cells[rowIndex, 2].GetStyle();
                                    style.Number = 4;   // 数字型
                                    // 小数位数
                                    style.Custom = digit switch
                                    {
                                        1 => "0.0",
                                        2 => "0.00",
                                        3 => "0.000",
                                        4 => "0.0000",
                                        _ => null,
                                    };
                                    // 设置
                                    cells[rowIndex, 2].SetStyle(style);
                                    cells[rowIndex, 3].SetStyle(style);

                                    // 计算当前点与上一个点的距离
                                    if (!double.IsNaN(prevX) && !double.IsNaN(prevY) && (rowIndex - 3) % 56 != 0)
                                    {
                                        string distance = Math.Sqrt(Math.Pow(x - prevX, 2) + Math.Pow(y - prevY, 2)).RoundWithFill(2);
                                        // 写入距离【第一行不写入】
                                        cells[rowIndex - 1, 1].Value = distance;
                                        if (j == mapPoints[i].Count - 1)
                                        {
                                            cells[rowIndex + 1, 1].Value = "";
                                        }
                                    }
                                    prevX = x;
                                    prevY = y;
                                    // 是否跨页
                                    if ((rowIndex - 47) % 56 == 0)
                                    {
                                        // 但是如果不是最后一个点，就回退
                                        if (pointIndex - lastRowCount != mapPoints[i].Count)
                                        {
                                            rowIndex += 12;
                                            // 点号回退1
                                            j--;
                                            pointIndex--;
                                        }
                                        else
                                        {
                                            rowIndex += 12;
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
        public static void SetPage(ArcGIS.Core.Geometry.Polygon polygon, Cells cells, string xmmc, string dkmc, string dkwz, string area_m, string area_ha, string lineLenth, string pointCount, string jsz, string jcz, string zbdw, string shdw, string zbrq, string sr)
        {

            // 获取要素的界址点总数
            var pointsCount = polygon.Points.Count;
            // 计算要生成的页数
            long pageCount = (int)Math.Ceiling((double)pointsCount / 23);

            if (pageCount > 1)
            {
                // 复制分页
                for (int i = 0; i < pageCount - 1; i++)
                {
                    cells.CopyRows(cells, 0, (i + 1) * 56, 55);
                }
            }

            // 设置单元格
            for (int i = 0; i < pageCount; i++)
            {
                // 标题
                cells[i * 56 + 0, 0].Value = $"{xmmc}（{dkmc}）\n\r界址点坐标成果表（{sr}）";
                // 页码
                cells[i * 56 + 1, 4].Value = $"第{i + 1}页  共 {pageCount} 页";
                // 地块位置
                cells[i * 56 + 1, 0].Value = $"地块位置：{dkwz}";
                // 界线总长
                cells[i * 56 + 49, 2].Value = lineLenth;
                // 界址点总数
                cells[i * 56 + 49, 4].Value = $"界址点总数：{pointCount}";
                // 地块面积
                cells[i * 56 + 50, 2].Value = area_m;
                cells[i * 56 + 50, 4].Value = area_ha;
                // 计算者
                cells[i * 56 + 51, 0].Value = $"计算者：{jsz}";
                // 检查者
                cells[i * 56 + 51, 2].Value = $"检查者：{jcz}";
                // 制表单位
                cells[i * 56 + 52, 1].Value = zbdw;
                // 审核单位
                cells[i * 56 + 52, 4].Value = shdw;
                // 制表日期
                cells[i * 56 + 54, 3].Value = $"日期：{zbrq}";
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/144554682";
            UITool.Link2Web(url);
        }

        private void combox_dkmc_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_dkmc);
        }

        private void combox_xmmc_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_xmmc);
        }

        private void combox_dkwz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_dkwz);
        }

        private void combox_jsz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_jsz);
        }

        private void combox_jcz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_jcz);
        }

        private void combox_zbdw_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_zbdw);
        }

        private void combox_shdw_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_shdw);
        }

        private void combox_zbrq_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_zbrq);
        }
    }
}
