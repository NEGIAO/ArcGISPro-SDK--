using ActiproSoftware.Windows.Shapes;
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
using Aspose.Words;
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
using Polygon = ArcGIS.Core.Geometry.Polygon;
using Range = Aspose.Cells.Range;

namespace CCTool.Scripts.CusTool4
{
    /// <summary>
    /// Interaction logic for QJDCExcel.xaml
    /// </summary>
    public partial class QJDCExcel : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "QJDCExcel";
        public QJDCExcel()
        {
            InitializeComponent();

            // 初始化参数选项
            textFolderPath.Text = BaseTool.ReadValueFromReg(toolSet, "excel_folder");
            cb_pdf.IsChecked = BaseTool.ReadValueFromReg(toolSet, "pdf").ToBool();
        }

        // 更新默认字段
        public async void UpdataField()
        {
            string ly = combox_fc.ComboxText();

            // 初始化参数选项
            await UITool.InitLayerFieldToComboxPlus(combox_zddm, ly, "ZDDM", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_qlr, ly, "TDSYZMC", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_zl, ly, ["ZL", "TDZL"], "string");
            await UITool.InitLayerFieldToComboxPlus(combox_fl, ly, "法人", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_sfz, ly, "法人SHZ", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_tfh, ly, "TFH", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_zdmj, ly, "ZDMJ", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_tdyt, ly, "TDYT", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_bdcdyh, ly, "BDCDYH", "string");

            await UITool.InitLayerFieldToComboxPlus(combox_txdz, ly, "SSXZCMC", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_frdh, ly, "法人SJH", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_bdch, ly, "原BDCZSH", "string");

            await UITool.InitLayerFieldToComboxPlus(combox_bz, ly, "ZDSZB", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_dz, ly, "ZDSZD", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_nz, ly, "ZDSZN", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_xz, ly, "ZDSZX", "string");

            await UITool.InitLayerFieldToComboxPlus(combox_nyd, ly, "农用地", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_gd, ly, "耕地", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_ld, ly, "林地", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_cd, ly, "草地", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_qt, ly, "其他", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_jsyd, ly, "建设用地", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_wlyd, ly, "未利用地", "float");
        }


        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "权属调查表(雨)";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textFolderPath.Text = UITool.OpenDialogFolder();
        }

        // 获取字段值，如无则为""
        private string GetFieldValue(Feature feature, string fieldName)
        {
            string result = "";

            if (fieldName is null || fieldName == "")
            {
                return result;
            }
            else
            {
                result = feature[fieldName].ToString();
                return result;
            }
        }


        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 参数获取
                string in_fc = combox_fc.ComboxText();
                string excel_folder = textFolderPath.Text;
                bool pdf = (bool)cb_pdf.IsChecked;

                string modelPath = textExcelModel.Text;

                string field_zddm = combox_zddm.ComboxText();
                string field_qlr = combox_qlr.ComboxText();
                string field_zl = combox_zl.ComboxText();
                string field_fl = combox_fl.ComboxText();
                string field_sfz = combox_sfz.ComboxText();
                string field_tfh = combox_tfh.ComboxText();
                string field_zdmj = combox_zdmj.ComboxText();
                string field_tdyt = combox_tdyt.ComboxText();
                string field_bdcdyh = combox_bdcdyh.ComboxText();

                string field_txdz = combox_txdz.ComboxText();
                string field_frdh = combox_frdh.ComboxText();
                string field_bdch = combox_bdch.ComboxText();

                string field_bz = combox_bz.ComboxText();
                string field_dz = combox_dz.ComboxText();
                string field_nz = combox_nz.ComboxText();
                string field_xz = combox_xz.ComboxText();

                string field_nyd = combox_nyd.ComboxText();
                string field_gd = combox_gd.ComboxText();
                string field_ld = combox_ld.ComboxText();
                string field_cd = combox_cd.ComboxText();
                string field_qt = combox_qt.ComboxText();
                string field_jsyd = combox_jsyd.ComboxText();
                string field_wlyd = combox_wlyd.ComboxText();

                // 判断参数是否选择完全
                if (in_fc == "" || excel_folder == "" || field_zddm == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "excel_folder", excel_folder);
                BaseTool.WriteValueToReg(toolSet, "pdf", pdf);

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

                    // 遍历面要素类中的所有要素
                    RowCursor cursor = featurelayer.TargetSelectCursor();
                    while (cursor.MoveNext())
                    {
                        using var feature = cursor.Current as Feature;

                        Polygon polygon = feature.GetShape() as Polygon;

                        // 获取参数
                        string oidField = in_fc.TargetIDFieldName();
                        string oid = feature[oidField].ToString();

                        string zddm = GetFieldValue(feature, field_zddm);

                        string qlr = GetFieldValue(feature, field_qlr);
                        string zl = GetFieldValue(feature, field_zl);
                        string fl = GetFieldValue(feature, field_fl);
                        string sfz = GetFieldValue(feature, field_sfz);
                        string tfh = GetFieldValue(feature, field_tfh);
                        string zdmj = GetFieldValue(feature, field_zdmj);
                        string tdyt = GetFieldValue(feature, field_tdyt);
                        string bdcdyh = GetFieldValue(feature, field_bdcdyh);

                        string txdz = GetFieldValue(feature, field_txdz);
                        string frdh = GetFieldValue(feature, field_frdh);
                        string bdch = GetFieldValue(feature, field_bdch);

                        string bz = GetFieldValue(feature, field_bz);
                        string dz = GetFieldValue(feature, field_dz);
                        string nz = GetFieldValue(feature, field_nz);
                        string xz = GetFieldValue(feature, field_xz);

                        string nyd = GetFieldValue(feature, field_nyd);
                        string gd = GetFieldValue(feature, field_gd);
                        string ld = GetFieldValue(feature, field_ld);
                        string cd = GetFieldValue(feature, field_cd);
                        string qt = GetFieldValue(feature, field_qt);
                        string jsyd = GetFieldValue(feature, field_jsyd);
                        string wlyd = GetFieldValue(feature, field_wlyd);

                        pw.AddMessageMiddle(20, $"处理要素：{oid} - 宗地码：{zddm}");

                        // 复制界址点Excel表，看下有没有模板文件
                        string excelPath = excel_folder + @$"\权籍调查表_{zddm}.xlsx";
                        if (modelPath == "")
                        {
                            DirTool.CopyResourceFile(@"CCTool.Data.Excel.杂七杂八.所有权模板王鹏飞0520.xlsx", excelPath);
                        }
                        else
                        {
                            File.Copy(modelPath, excelPath);
                        }

                        // 打开工作薄
                        Workbook wb = ExcelTool.OpenWorkbook(excelPath);

                        // 处理sheet_不动产登记申请书1
                        Worksheet ws_sqs1 = wb.Worksheets["不动产登记申请书1"];
                        Cells cells_sqs1 = ws_sqs1.Cells;
                        cells_sqs1["E8"].Value = $"{qlr}";
                        cells_sqs1["E10"].Value = $"{txdz}";
                        cells_sqs1["H7"].Value = $"{txdz}";

                        cells_sqs1["E21"].Value = $"{zl}";
                        cells_sqs1["E22"].Value = $"{zddm}W00000000";
                        cells_sqs1["E23"].Value = $"{zdmj}";

                        // 处理sheet_不动产权籍调查表 
                        Worksheet ws_fm = wb.Worksheets["不动产权籍调查表"];
                        Cells cells_fm = ws_fm.Cells;
                        string text = cells_fm["A2"].StringValue;
                        cells_fm["A2"].Value = text.Replace("宗地 / 宗海代码 ：", $"宗地 / 宗海代码 ：{zddm}");

                        // 处理sheet_地籍调查表 
                        Worksheet ws_dj = wb.Worksheets["地籍调查表"];
                        Cells cells_dj = ws_dj.Cells;
                        cells_dj["C3"].Value = $"{qlr}";
                        cells_dj["C9"].Value = $"{zl}";

                        cells_dj["C16"].Value = $"{zddm[6..]}";
                        cells_dj["J16"].Value = $"{zddm}";
                        cells_dj["C17"].Value = $"{zddm}W00000000";
                        cells_dj["E19"].Value = $"{tfh}";

                        cells_dj["C20"].Value = $"北：{bz}";
                        cells_dj["C21"].Value = $"东：{dz}";
                        cells_dj["C22"].Value = $"南：{nz}";
                        cells_dj["C23"].Value = $"西：{xz}";

                        //cells_dj["N26"].Value = $"{tdyt}";
                        cells_dj["H27"].Value = $"{zdmj}";

                        // 处理sheet_调查审核表
                        Worksheet ws_sh = wb.Worksheets["调查审核表"];
                        Cells cells_sh = ws_sh.Cells;
                        cells_sh["B4"].Value = $"测量前经检查，对本宗地共 {polygon.PointCount - 1}个界址点，全部使用全站仪采用解析法测定界址点坐标。";

                        // 处理sheet_集体土地所有权宗地分类面积调查表
                        Worksheet ws_syq = wb.Worksheets["集体土地所有权宗地分类面积调查表"];
                        Cells cells_syq = ws_syq.Cells;
                        cells_syq["E3"].Value = $"{qlr}";
                        cells_syq["E4"].Value = $"{zddm}";
                        cells_syq["E5"].Value = $"{zddm}W00000000";

                        cells_syq["F6"].Value = $"{nyd}";
                        cells_syq["F7"].Value = $"{gd}";
                        cells_syq["F8"].Value = $"{ld}";
                        cells_syq["F9"].Value = $"{cd}";
                        cells_syq["F10"].Value = $"{qt}";
                        cells_syq["F11"].Value = $"{jsyd}";
                        cells_syq["F12"].Value = $"{wlyd}";


                        //// 处理sheet_宗地图
                        //Worksheet ws_zd = wb.Worksheets["宗地图"];
                        //Cells cells_zd = ws_zd.Cells;
                        //cells_zd["A2"].Value = $"宗地代码：{zddm}";
                        //cells_zd["B2"].Value = $"土地权利人：{qlr}";
                        //cells_zd["A3"].Value = $" 所在图幅号：{tfh}";
                        //cells_zd["B3"].Value = $"宗地面积：{zdmj}";

                        // 保存
                        wb.Save(excelPath);
                        wb.Dispose();

                        // 界址标示表
                        JZBSB($@"{excelPath}\界址标示表$", polygon);

                        // 界址点成果表
                        JZDB($@"{excelPath}\界址点成果表$", polygon, qlr, zddm, zdmj);

                        // 界址说明表
                        JZSM($@"{excelPath}\界址说明表$", polygon, qlr);

                        // 导出PDF
                        if (pdf)
                        {
                            ExcelTool.ImportToPDF(excelPath, excel_folder + @$"\权籍调查表_{zddm}.pdf");
                        }

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


        // 界址标示表
        private void JZBSB(string excelPath, Polygon polygon)
        {
            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 处理sheet
            Worksheet ws = wb.Worksheets[sheetIndex];
            Cells cells = ws.Cells;

            // 获取面要素的所有折点【按西北角起始，顺时针重排】
            List<List<MapPoint>> mapPoints = polygon.ReshotMapPoint();

            // 设置分页信息
            SetPage(mapPoints, cells);


            int index = 3;   // 起始行
            int pointIndex = 1;  // 起始点序号
            int lastRowCount = 0;  // 上一轮的行数
            int ptDigit = 2;

            for (int i = 0; i < mapPoints.Count; i++)
            {
                // 输出折点的XY值和距离到Excel表
                double prevX = double.NaN;
                double prevY = double.NaN;

                for (int j = 0; j < mapPoints[i].Count; j++)
                {
                    // 写入点号
                    if (index > 3)
                    {
                        if (pointIndex - lastRowCount == mapPoints[i].Count)    // 找到当前环的最后一点
                        {
                            cells[index , 0].Value = $"J{lastRowCount + 1 - i}";
                        }
                        else
                        {
                            cells[index, 0].Value = $"J{pointIndex - i}";
                        }
                    }
                    else
                    {
                        cells[index, 0].Value = $"J1";
                    }

                    double x = Math.Round(mapPoints[i][j].X, ptDigit);
                    double y = Math.Round(mapPoints[i][j].Y, ptDigit);

                    // 设置单元格为数字型，小数位数
                    Aspose.Cells.Style style = cells[index - 1, 6].GetStyle();
                    style.Number = 4;   // 数字型
                    style.Custom = ptDigit switch
                    {
                        1 => "0.0",
                        2 => "0.00",
                        3 => "0.000",
                        4 => "0.0000",
                        _ => null,
                    };
                    // 设置
                    cells[index - 1, 6].SetStyle(style);
                    cells[index, 6].SetStyle(style);

                    // 计算当前点与上一个点的距离
                    if (!double.IsNaN(prevX) && !double.IsNaN(prevY) && index > 3)
                    {
                        string distance = Math.Sqrt(Math.Pow(x - prevX, 2) + Math.Pow(y - prevY, 2)).RoundWithFill(ptDigit);
                        // 写入距离【第一行不写入】
                        if (index == 3)
                        {
                            continue;
                        }
                        else if (index == 42)
                        {
                            cells[index-5, 6].Value = distance;
                            cells[index, 13].Value = "";
                            cells[index , 15].Value = "";
                        }
                        else if ((index-40) % 42 == 0 && (index - 40) / 42 > 0)
                        {
                            cells[index - 7, 6].Value = distance;
                            cells[index , 13].Value = "";
                            cells[index, 15].Value = "";
                        }
                        else
                        {
                            cells[index-1, 6].Value = distance;
                        }

                        
                        if (j == mapPoints[i].Count - 1)
                        {
                            cells[index, 6].Value = "";
                        }
                    }
                    prevX = x;
                    prevY = y;

                    // 行变化
                    if (index == 3)
                    {
                        index += 1;
                    }
                    else if (index == 36)
                    {
                        index +=6;
                    }
                    else if (index % 40 == 34 && index / 40 >0)
                    {
                        index += 8;
                    }
                    else
                    {
                        index += 2;
                    }

                    pointIndex++;
                }
                lastRowCount += mapPoints[i].Count;
            }

            // 删除多余行,100是随意的
            cells.DeleteRows(index, 100);
            cells[index-1, 13].Value = "";
            cells[index-1, 15].Value = "";


            Aspose.Cells.Style style4;

            if (index == 44)     // 第一次跨页第一行
            {
                style4 = cells[index - 7, 6].GetStyle();
                cells[index - 2, 6].SetStyle(style4);
            }

            else if (index % 40 == 44 && index / 40 > 1)     // 后续跨页第一行
            {
                style4 = cells[index - 9, 6].GetStyle();
                cells[index - 2, 6].SetStyle(style4);
            }
            else if (index % 40 == 42 )     // 尾行
            {
                cells[index-5, 13].Value = "";
                cells[index-5, 15].Value = "";

                cells.DeleteRows(index-4, 100);
            }
            else
            {
                style4 = cells[index - 3, 6].GetStyle();
                cells[index - 2, 6].SetStyle(style4);
            }

            

            //// 合并单元格
            //for (int k = 4; k < index - 2; k += 2)
            //{
            //    for (int j = 0; j < 6; j++)
            //    {
            //        cells.Merge(k, j, 2, 1);
            //    }
            //}

            //for (int k = 3; k < index - 3; k += 2)
            //{
            //    for (int j = 6; j < 19; j++)
            //    {
            //        cells.Merge(k, j, 2, 1);
            //    }
            //}


            // 保存
            wb.Save(excelFile);
            wb.Dispose();

        }

        // 界址点表
        private void JZDB(string excelPath, Polygon polygon, string qlr, string zddm, string mj)
        {
            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 处理sheet
            Worksheet ws = wb.Worksheets[sheetIndex];
            Cells cells = ws.Cells;


            // 设置分页信息
            SetPage(polygon, cells, qlr, zddm, mj.ToDouble());

            if (polygon != null)
            {
                // 获取面要素的所有折点【按西北角起始，顺时针重排】
                List<List<MapPoint>> mapPoints = polygon.ReshotMapPoint();

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

                        if (pointIndex - lastRowCount == mapPoints[i].Count)    // 找到当前环的最后一点
                        {
                            cells[rowIndex, 1].Value = $"J{lastRowCount + 1 - i}";
                        }
                        else
                        {
                            cells[rowIndex, 1].Value = $"J{pointIndex - i}";
                        }

                        // 写入序号
                        cells[rowIndex, 0].Value = $"{pointIndex}";

                        double x = Math.Round(mapPoints[i][j].X, 3);
                        double y = Math.Round(mapPoints[i][j].Y, 3);
                        // 写入折点的XY值
                        cells[rowIndex, 5].Value = x;
                        cells[rowIndex, 4].Value = y;
                        // 设置单元格为数字型，小数位数
                        Aspose.Cells.Style style = cells[rowIndex, 4].GetStyle();
                        style.Number = 4;   // 数字型
                                            // 小数位数
                        style.Custom = 3 switch
                        {
                            1 => "0.0",
                            2 => "0.00",
                            3 => "0.000",
                            4 => "0.0000",
                            _ => null,
                        };
                        // 设置
                        cells[rowIndex, 5].SetStyle(style);
                        cells[rowIndex, 4].SetStyle(style);

                        // 计算当前点与上一个点的距离
                        if (!double.IsNaN(prevX) && !double.IsNaN(prevY) && rowIndex > 8)
                        {
                            string distance = Math.Sqrt(Math.Pow(x - prevX, 2) + Math.Pow(y - prevY, 2)).RoundWithFill(2);
                            // 写入距离【第一行不写入】
                            cells[rowIndex - 1, 6].Value = distance;
                            if (j == mapPoints[i].Count - 1)
                            {
                                cells[rowIndex + 1, 6].Value = "";
                            }
                        }
                        prevX = x;
                        prevY = y;
                        // 是否跨页
                        if ((rowIndex + 3) % 57 == 0)
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
            wb.Save(excelFile);
            wb.Dispose();

        }

        // 生成分页，并设置分页信息
        public static void SetPage(Polygon geometry, Cells cells, string QLR, string ZDDM, double ZDMJ)
        {
            // 获取要素的界址点总数
            var pointsCount = geometry.Points.Count;
            // 计算要生成的页数
            long pageCount = (int)Math.Ceiling((double)pointsCount / 24);

            if (pageCount > 1)  // 多于1页
            {
                // 复制分页，前两页已有不用复制
                for (int i = 1; i < pageCount; i++)
                {
                    cells.CopyRows(cells, 0, i * 57, 57);
                }
            }
            // 设置单元格
            for (int i = 0; i < pageCount; i++)
            {
                // 填写页码
                cells[i * 57, 6].Value = $"第 {i + 1} 页";
                cells[i * 57 + 1, 6].Value = $"共 {pageCount} 页";
                // 宗地号
                cells[i * 57 + 2, 2].Value = ZDDM;
                // 权利人
                cells[i * 57 + 3, 2].Value = QLR;
                // 面积
                cells[i * 57 + 4, 3].Value = ZDMJ;
            }
        }

        // 界址点表
        private void JZSM(string excelPath, Polygon polygon, string QLR)
        {
            // 点位说明
            string dysm = "";
            // 走向说明
            string zxsm = "";

            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 处理sheet
            Worksheet ws = wb.Worksheets[sheetIndex];
            Cells cells = ws.Cells;


            if (polygon != null)
            {
                // 获取面要素的所有折点【按西北角起始，顺时针重排】
                List<List<MapPoint>> mapPoints = polygon.ReshotMapPoint();
                // 面的中心点坐标
                MapPoint centerPoint = GeometryEngine.Instance.Centroid(polygon);

                int pointIndex = 1;  // 起始点序号
                int lastRowCount = 0;  // 上一轮的行数

                for (int i = 0; i < mapPoints.Count; i++)
                {
                    // 输出折点的XY值和距离到Excel表
                    double prevX = double.NaN;
                    double prevY = double.NaN;

                    // 上一个点的点号
                    string lastJP = "";

                    for (int j = 0; j < mapPoints[i].Count; j++)
                    {
                        // 当前点的点号
                        string initJP = "";

                        // 计算界址点和面的质心的方位
                        string direction = GeometryTool.Get4Direction(centerPoint, mapPoints[i][j]);

                        if (pointIndex - lastRowCount == mapPoints[i].Count)    // 找到当前环的最后一点
                        {
                            initJP = $"J{lastRowCount + 1 - i}";
                        }
                        else
                        {
                            initJP = $"J{pointIndex - i}";
                        }
                        // 宗地说明
                        dysm += $"{initJP}位于{QLR}宗地{direction}；";

                        double x = Math.Round(mapPoints[i][j].X, 3);
                        double y = Math.Round(mapPoints[i][j].Y, 3);

                        if (!double.IsNaN(prevX) && !double.IsNaN(prevY))
                        {
                            // 计算当前点与上一个点的距离
                            string distance = Math.Sqrt(Math.Pow(x - prevX, 2) + Math.Pow(y - prevY, 2)).RoundWithFill(2);
                            // 方位角
                            string direction2 = GeometryTool.Get8Direction(prevX, prevY, x, y);

                            // 走向说明
                            zxsm += $"{lastJP}沿两点连线(中),{direction2}方向{distance}米走向至{initJP}；";
                        }
                        prevX = x;
                        prevY = y;
                        // 当前点赋值给前一个点
                        lastJP = initJP;
                        pointIndex++;
                    }

                    lastRowCount += mapPoints[i].Count;
                }
            }

            // 写入
            cells["B2"].Value = dysm;
            cells["B3"].Value = zxsm;

            // 保存
            wb.Save(excelFile);
            wb.Dispose();

        }

        private void combox_zddm_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_zddm);
        }

        private void combox_qlr_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_qlr);
        }

        private void combox_zl_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_zl);
        }

        private void combox_fl_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_fl);
        }

        private void combox_sfz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_sfz);
        }

        private void combox_tfh_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_tfh);
        }

        private void combox_zdmj_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_zdmj);
        }

        private void combox_tdyt_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_tdyt);
        }

        private void combox_bdcdyh_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_bdcdyh);
        }

        private void combox_bz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_bz);
        }

        private void combox_dz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_dz);
        }

        private void combox_nz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_nz);
        }

        private void combox_xz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_xz);
        }

        private void combox_nyd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_nyd);
        }

        private void combox_gd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_gd);
        }

        private void combox_ld_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_ld);
        }

        private void combox_cd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_cd);
        }

        private void combox_qt_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_qt);
        }

        private void combox_jsyd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_jsyd);
        }

        private void combox_wlyd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_wlyd);
        }

        private void combox_bdch_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_bdch);
        }

        private void combox_txdz_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_txdz);
        }

        private void combox_frdh_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_frdh);
        }

        private void combox_fc_DropClose(object sender, EventArgs e)
        {
            // 更新默认字段
            UpdataField();
        }

        private void openModelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelModel.Text = UITool.OpenDialogExcel();
        }

        // 生成分页，并设置分页信息
        public static void SetPage(List<List<MapPoint>> mapPoints, Cells cells)
        {
            // 获取要素的界址点总数
            long pointsCount = 0;
            foreach (var pts in mapPoints)
            {
                pointsCount += pts.Count;
            }

            // 计算要生成的页数
            long pageCount = (int)Math.Ceiling((double)(pointsCount - 18) / 17) + 1;

            if (pageCount == 1)   // 只有一页，就删第二页
            {
                cells.DeleteRows(38, 50);
            }
            else if (pageCount > 2)  // 多于2页
            {
                // 复制分页，前两页已有不用复制
                for (int i = 0; i < pageCount - 2; i++)
                {
                    cells.CopyRows(cells, 38, 78 + i * 40, 40);
                }
            }
        }
    }
}
