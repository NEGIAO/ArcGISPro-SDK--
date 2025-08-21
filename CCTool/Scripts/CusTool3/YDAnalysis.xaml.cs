using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Row = ArcGIS.Core.Data.Row;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for YDAnalysis.xaml
    /// </summary>
    public partial class YDAnalysis : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "YDAnalysis";
        public YDAnalysis()
        {
            InitializeComponent();

            //UITool.InitFeatureLayerToComboxPlus(combox_pd, "已批农转");
            //UITool.InitFeatureLayerToComboxPlus(combox_gd, "已供应");

            //UITool.InitFeatureLayerToComboxPlus(combox_kfbj, "开发边界");
            //UITool.InitFeatureLayerToComboxPlus(combox_xzq, "丰山镇");

            //UITool.InitFeatureLayerToComboxPlus(combox_yd, "新增数据");

            // 初始化参数选项
            textOutFCPath.Text = BaseTool.ReadValueFromReg(toolSet, "outFCPath");
            textExcelPath.Text = BaseTool.ReadValueFromReg(toolSet, "excelPath");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "批供地分析";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                string defGDB = Project.Current.DefaultGeodatabasePath;

                // 获取参数
                string pd = combox_pd.ComboxText();
                string gd = combox_gd.ComboxText();
                string kfbj = combox_kfbj.ComboxText();
                string xzq = combox_xzq.ComboxText();
                string yd = combox_yd.ComboxText();

                string outFCPath = textOutFCPath.Text;
                string excelPath = textExcelPath.Text;


                // 判断参数是否选择完全
                if (pd == "" || gd == "" || kfbj == "" || xzq == "" || yd == "" || outFCPath == "" || excelPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "outFCPath", outFCPath);
                BaseTool.WriteValueToReg(toolSet, "excelPath", excelPath);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(yd, xzq);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10, "获取项目指标");
                    // 融合
                    string mcField = "XMMC";
                    string xzqField = "XZQMC";
                    string yd_dissolve = @$"{defGDB}\yd_dissolve";
                    Arcpy.Dissolve(yd, yd_dissolve, mcField);
                    // 标识
                    string yd_identity = @$"{defGDB}\yd_identity";
                    Arcpy.Identity(yd_dissolve, xzq, yd_identity);
                    // 获取面积
                    List<string> strings = new List<string>() { xzqField, mcField };
                    Dictionary<List<string>, double> dkDict = GisTool.GetDictListFromPathDouble(yd_identity, strings, "shape_area");

                    pw.AddMessageMiddle(20, "获取批供地指标");
                    // 裁剪
                    string pd_clip = @$"{defGDB}\pd_clip";
                    string gd_clip = @$"{defGDB}\gd_clip";
                    string wg = @$"{defGDB}\wg";
                    Arcpy.Clip(yd, pd, pd_clip);
                    Arcpy.Clip(yd, gd, gd_clip);
                    // 批而未供
                    Arcpy.Erase(pd_clip, gd_clip, wg);
                    // 添加计算标记
                    string bjField = "类型标记";
                    Arcpy.AddField(pd_clip, bjField, "TEXT");
                    Arcpy.AddField(gd_clip, bjField, "TEXT");
                    Arcpy.AddField(wg, bjField, "TEXT");
                    Arcpy.CalculateField(pd_clip, bjField, "'已批用地'");
                    Arcpy.CalculateField(gd_clip, bjField, "'已供用地'");
                    Arcpy.CalculateField(wg, bjField, "'批而未供用地'");
                    // 追加
                    Arcpy.Append($"{pd_clip};{gd_clip}", wg);
                    // 开发边界裁剪
                    string wg_clip = @$"{defGDB}\wg_clip";
                    Arcpy.Clip(wg, kfbj, wg_clip);
                    // 删除字段
                    Arcpy.DeleteField(wg_clip, new List<string>() { mcField, bjField }, "KEEP_FIELDS");
                    // 标识
                    string pgd = @$"{defGDB}\pgd";
                    Arcpy.Identity(wg_clip, xzq, pgd);
                    // 筛选批而未供用地
                    Arcpy.Select(pgd, $@"{outFCPath}\批而未供用地", "类型标记 = '批而未供用地'");
                    Arcpy.Select(pgd, $@"{outFCPath}\已批用地", "类型标记 = '已批用地'");
                    Arcpy.Select(pgd, $@"{outFCPath}\已供用地", "类型标记 = '已供用地'");
                    // 汇总
                    string pgd_sta = @$"{defGDB}\pgd_sta";
                    Arcpy.Statistics(pgd, pgd_sta, $"shape_area SUM", $"{xzqField};{mcField};{bjField}");
                    // 获取面积
                    List<string> strings2 = new List<string>() { xzqField, mcField, bjField };
                    Dictionary<List<string>, double> pgdDict = GisTool.GetDictListFromPathDouble(pgd_sta, strings2, "SUM_Shape_Area");

                    pw.AddMessageMiddle(30, "写入Excel");
                    // 复制excel
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.规划.汇总.xlsx", excelPath);
                    // 写入地块面积
                    int index = WriteExcelDK(excelPath, dkDict);
                    WriteExcelPGD(excelPath, pgdDict);

                    // 统计合计值
                    for (int i = 3; i < 7; i++)
                    {
                        ExcelTool.StatisticsColCell(excelPath, i, 1, index - 1, index);
                    }

                    pw.AddMessageMiddle(10, "清除中间数据");

                    List<string> paths = new List<string>() { yd_dissolve, yd_identity, pd_clip , gd_clip , wg , wg_clip , pgd , pgd_sta };
                    foreach (string path in paths)
                    {
                        Arcpy.Delect(path);
                    }

                });

                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private int WriteExcelDK(string excelPath, Dictionary<List<string>, double> dkDict)
        {
            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            Cells cells = sheet.Cells;

            int index = 1;
            foreach (var dk in dkDict)
            {
                if (index > 1)
                {
                    cells.CopyRow(cells, 1, index);
                }

                string xzqmc = dk.Key[0];
                string xmmc = dk.Key[1];
                double mj = dk.Value / 10000 * 15;

                cells[index, 0].Value = index;
                cells[index, 1].Value = xzqmc;
                cells[index, 2].Value = xmmc;
                cells[index, 3].Value = mj;

                index++;
            }
            // 复制合计行
            cells.CopyRow(cells, 1, index);
            cells.Merge(index, 0, 1, 3);
            cells[index, 0].Value = "合计";
            // 保存
            wb.Save(excelFile);
            wb.Dispose();

            return index;
        }

        private void WriteExcelPGD(string excelPath, Dictionary<List<string>, double> pgdDict)
        {
            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            Cells cells = sheet.Cells;

            foreach (var dk in pgdDict)
            {
                string xzqmc = dk.Key[0];
                string xmmc = dk.Key[1];
                string bj = dk.Key[2];
                double mj = dk.Value / 10000 * 15;

                for (int i = 1; i <= cells.MaxDataRow; i++)
                {
                    string xzqmc_ori = cells[i, 1].StringValue;
                    string xmmc_ori = cells[i, 2].StringValue;
                    // 对比，如果一致就写入
                    if (xzqmc == xzqmc_ori && xmmc == xmmc_ori)
                    {
                        if (bj == "已供用地") { cells[i, 4].Value = mj; }
                        else if (bj == "已批用地") { cells[i, 5].Value = mj; }
                        else if (bj == "批而未供用地") { cells[i, 6].Value = mj; }
                    }
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }


        private List<string> CheckData(string yd, string xzq)
        {
            List<string> result = new List<string>();

            if (yd != "")
            {
                // 检查字段值是否符合要求
                string result_value = CheckTool.CheckFieldValueSpace(yd, "XMMC");
                if (result_value != "")
                {
                    result.Add(result_value);
                }
            }

            if (xzq != "")
            {
                // 检查字段值是否符合要求
                string result_value = CheckTool.IsHaveFieldInLayer(xzq, "XZQMC");
                if (result_value != "")
                {
                    result.Add(result_value);
                }
            }
            return result;
        }

        private void combox_pd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_pd);
        }

        private void combox_gd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_gd);
        }

        private void combox_kfbj_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_kfbj);
        }

        private void combox_xzq_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_xzq);
        }

        private void combox_yd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_yd);
        }

        private void openOutFCButton_Click(object sender, RoutedEventArgs e)
        {
            textOutFCPath.Text = UITool.OpenDialogGDB();
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.SaveDialogExcel();
        }
    }
}
