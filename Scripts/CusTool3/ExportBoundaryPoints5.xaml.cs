using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Wordprocessing;
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


namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for ExportBoundaryPoints5.xaml
    /// </summary>
    public partial class ExportBoundaryPoints5 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        string toolSet = "ExportBoundaryPoints5";
        public ExportBoundaryPoints5()
        {
            InitializeComponent();

            combox_xyDigit.Items.Add("1");
            combox_xyDigit.Items.Add("2");
            combox_xyDigit.Items.Add("3");
            combox_xyDigit.Items.Add("4");
            combox_xyDigit.Items.Add("5");
            combox_xyDigit.Items.Add("6");
            combox_xyDigit.SelectedIndex = 1;

            combox_mDigit.Items.Add("1");
            combox_mDigit.Items.Add("2");
            combox_mDigit.Items.Add("3");
            combox_mDigit.Items.Add("4");
            combox_mDigit.Items.Add("5");
            combox_mDigit.Items.Add("6");
            combox_mDigit.SelectedIndex = 1;

            combox_haDigit.Items.Add("1");
            combox_haDigit.Items.Add("2");
            combox_haDigit.Items.Add("3");
            combox_haDigit.Items.Add("4");
            combox_haDigit.Items.Add("5");
            combox_haDigit.Items.Add("6");
            combox_haDigit.SelectedIndex = 3;


            // 读取参数
            textFolderPath.Text = BaseTool.ReadValueFromReg(toolSet, "excelFolder");
            txt_js.Text = BaseTool.ReadValueFromReg(toolSet, "js");
            txt_jc.Text = BaseTool.ReadValueFromReg(toolSet, "jc");
            txt_time.Text = BaseTool.ReadValueFromReg(toolSet, "time");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "界址点导出Excel(子弹)";


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
                string fc = combox_fc.ComboxText();
                string excelFolder = textFolderPath.Text;

                string mcField = combox_dkmc.ComboxText();

                string js = txt_js.Text;
                string jc = txt_jc.Text;
                string time = txt_time.Text;

                int xyDigit = combox_xyDigit.Text.ToInt();
                int mDigit = combox_mDigit.Text.ToInt();
                int haDigit = combox_haDigit.Text.ToInt();

                bool isDH = (bool)cb_DH.IsChecked;

                // 记录参数
                BaseTool.WriteValueToReg(toolSet, "excelFolder", excelFolder);
                BaseTool.WriteValueToReg(toolSet, "js", js);
                BaseTool.WriteValueToReg(toolSet, "jc", jc);
                BaseTool.WriteValueToReg(toolSet, "time", time);

                // 判断参数是否选择完全
                if (fc == "" || mcField == "" || excelFolder == "")
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

                    // 点号标记
                    int dh = 1;

                    // 获取目标FeatureLayer
                    FeatureLayer featurelayer = fc.TargetFeatureLayer();

                    // 遍历面要素类中的所有要素
                    RowCursor cursor = featurelayer.Search();
                    while (cursor.MoveNext())
                    {
                        // 点号是否重置
                        if (!isDH)
                        {
                            dh = 1;
                        }
                        
                        using var feature = cursor.Current as Feature;
                        // 获取要素的几何
                        Polygon polygon = feature.GetShape() as Polygon;
                        // 名称
                        string dkmc = feature[mcField]?.ToString();   // 地块名称

                        pw.AddMessageMiddle(10, $"      处理要素：{dkmc}", Brushes.Gray);

                        // 面积_平方米，公顷
                        string area_m = polygon.Area.RoundWithFill(mDigit) + "平方米";
                        string area_ha = (polygon.Area / 10000).RoundWithFill(haDigit) + "公顷";

                        // 复制界址点Excel表
                        string excelPath = @$"{excelFolder}\{dkmc}_界址点成果表.xlsx";
                        DirTool.CopyResourceFile(@"CCTool.Data.Excel.杂七杂八.界址点成果表模板样式.xlsx", excelPath);
                        // 获取工作薄、工作表
                        string excelFile = ExcelTool.GetPath(excelPath);
                        int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
                        // 打开工作薄
                        Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                        // 打开工作表
                        Worksheet worksheet = wb.Worksheets[sheetIndex];
                        // 获取Cells
                        Cells cells = worksheet.Cells;

                        // 设置信息
                        string originStr = "计算者:张三                  检查者: 李四                    2025年4月19日";
                        cells[11, 0].Value = originStr.Replace("张三", js).Replace("李四", jc).Replace("2025年4月19日", time);
                        cells[0, 0].Value = $"CGCS2000坐标系 {dkmc}界址点坐标成表";

                        if (polygon != null)
                        {
                            // 面积
                            string originMJ = "面积 =  100.00平方米  约0.01公顷";
                            cells[10, 0].Value = originMJ.Replace("100.00平方米", area_m).Replace("0.01公顷", area_ha);

                            // 修改最后一个点号
                            cells[8, 0].Value = $"J{dh}";

                            // 获取面要素的所有折点【按西北角起始，顺时针重排】
                            List<List<MapPoint>> mapPoints = polygon.ReshotMapPoint();

                            int index = 2;   // 起始行

                            for (int i = 0; i < mapPoints.Count; i++)
                            {
                                // 输出折点的XY值和距离到Excel表
                                double prevX = double.NaN;
                                double prevY = double.NaN;

                                for (int j = 0; j < mapPoints[i].Count - 1; j++)
                                {
                                    // 复制行
                                    if (index > 6)
                                    {
                                        cells.InsertRows(index, 2);
                                        cells.CopyRows(cells, 4, index, 2);
                                        cells.CopyRows(cells, 4, index, 2);
                                    }

                                    double x = Math.Round(mapPoints[i][j].X, xyDigit);
                                    double y = Math.Round(mapPoints[i][j].Y, xyDigit);

                                    // 写入折点的XY值
                                    cells[index, 2].Value = y;
                                    cells[index, 3].Value = x;

                                    // 写入点号
                                    cells[index, 0].Value = $"J{dh}";

                                    // 第一个点的时候，末点给填上
                                    if (j == 0)
                                    {
                                        cells[8, 2].Value = y;
                                        cells[8, 3].Value = x;
                                    }

                                    // 设置单元格为数字型，小数位数
                                    Aspose.Cells.Style style = cells[index, 2].GetStyle();
                                    style.Number = 4;   // 数字型
                                    // 小数位数
                                    style.Custom = xyDigit switch
                                    {
                                        1 => "0.0",
                                        2 => "0.00",
                                        3 => "0.000",
                                        4 => "0.0000",
                                        _ => null,
                                    };
                                    // 设置
                                    cells[index, 2].SetStyle(style);
                                    cells[index, 3].SetStyle(style);
                                    cells[index + 2, 2].SetStyle(style);
                                    cells[index + 2, 3].SetStyle(style);

                                    // 计算当前点与上一个点的距离
                                    if (!double.IsNaN(prevX) && !double.IsNaN(prevY))
                                    {
                                        string distance = Math.Sqrt(Math.Pow(x - prevX, 2) + Math.Pow(y - prevY, 2)).RoundWithFill(2);
                                        // 写入距离【第一行不写入】
                                        cells[index - 1, 1].Value = distance;
                                    }
                                    prevX = x;
                                    prevY = y;

                                    index += 2;
                                    dh++;
                                }
                            }

                            // 单元格处理
                            cells.UnMerge(7, 1, index - 7 + 1, 1);
                            cells.UnMerge(7, 4, index - 7 + 1, 1);

                            for (int i = 7; i < index; i += 2)
                            {
                                cells.Merge(i, 1, 2, 1);
                                cells.Merge(i, 4, 2, 1);
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


        private void combox_dkmc_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_dkmc);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147430378?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }
    }
}
