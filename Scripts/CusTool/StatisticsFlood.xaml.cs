using ArcGIS.Core.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Vml;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
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

namespace CCTool.Scripts.CusTool
{
    /// <summary>
    /// Interaction logic for StatisticsFlood.xaml
    /// </summary>
    public partial class StatisticsFlood : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public StatisticsFlood()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "洪水四线分析";

        private void combox_sd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_sd);
            textExcelPath.Text = Project.Current.HomeFolderPath + @"\洪水四线分析.xlsx";
        }

        private void combox_df_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_df);
        }

        private void combox_lsx_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_lsx);
        }

        private void combox_bcx_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_bcx);
        }

        private void combox_xhx_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_xhx);
        }

        private void combox_tcx_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_tcx);
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.SaveDialogExcel();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                string def_gdb = Project.Current.DefaultGeodatabasePath;
                string def_path = Project.Current.HomeFolderPath;
                // 获取分析要素
                string sd = combox_sd.ComboxText();
                string df = combox_df.ComboxText();

                string lsx = combox_lsx.ComboxText();
                string bcx = combox_bcx.ComboxText();
                string xhx = combox_xhx.ComboxText();
                string tcx = combox_tcx.ComboxText();

                // 输出路径
                string excel_path = textExcelPath.Text;

                // 判断参数是否选择完全
                if (sd == "" || df == "" || excel_path == "")
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
                    List<string> lines = new List<string>() { sd, df, lsx, bcx, xhx, tcx };
                    // 检查数据
                    List<string> errs = CheckData(lines, sd);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    // 复制嵌入资源中的Excel文件
                    string excel_mapper = $@"{def_path}\三调用地自转换_洪水.xlsx";
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.三调用地自转换_洪水.xlsx", excel_mapper);
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.【模板】洪水三上数据地类统计表.xlsx", excel_path);

                    pw.AddMessageMiddle(10, "三调数据预处理");

                    // 裁剪三调
                    string sd_clip = $@"{def_gdb}\sd_clip";
                    Arcpy.Clip(sd, df, sd_clip);
                    // 用地分类
                    Dictionary<string, string> MapDict = ExcelTool.GetDictFromExcel(excel_mapper);

                    pw.AddMessageMiddle(20, "地类划分");

                    string gnField = "分类";
                    Arcpy.AddField(sd_clip, gnField, "TEXT");
                    ComboTool.AttributeMapper(sd_clip, "DLBM", gnField, excel_mapper + @"\sheet1$");

                    // 4个范围分别裁剪统计范围，并汇总
                    pw.AddMessageMiddle(10, "裁剪汇总_临水线");
                    if (lsx != "")
                    {
                        ClipStatistics(lsx, df, sd_clip, gnField, excel_path, 2);
                    }
                    pw.AddMessageMiddle(10, "裁剪汇总_淹没补偿线");
                    if (bcx != "")
                    {
                        ClipStatistics(bcx, df, sd_clip, gnField, excel_path, 5);
                    }
                    pw.AddMessageMiddle(10, "裁剪汇总_秋汛行洪水边线");
                    if (xhx != "")
                    {
                        ClipStatistics(xhx, df, sd_clip, gnField, excel_path, 8);
                    }
                    pw.AddMessageMiddle(10, "裁剪汇总_滩槽分界线");
                    if (tcx != "")
                    {
                        ClipStatistics(tcx, df, sd_clip, gnField, excel_path, 11);
                    }

                    // 删除中间数据
                    Arcpy.Delect(sd_clip);
                    Arcpy.Delect($@"{def_gdb}\line_eraseArea");
                    Arcpy.Delect($@"{def_gdb}\line_clip");
                    Arcpy.Delect($@"{def_gdb}\line_erase");
                    Arcpy.Delect($@"{def_gdb}\line_clipTable");
                    Arcpy.Delect($@"{def_gdb}\line_eraseTable");

                    File.Delete(excel_mapper);

                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "";
            UITool.Link2Web(url);
        }

        // 分区裁剪汇总
        public void ClipStatistics(string line, string df, string sd_clip, string gnField, string excel_path, int col)
        {
            string def_gdb = Project.Current.DefaultGeodatabasePath;

            string erase_area = $@"{def_gdb}\line_eraseArea";
            string clip_fc = $@"{def_gdb}\line_clip";
            string erase_fc = $@"{def_gdb}\line_erase";

            string clip_table = $@"{def_gdb}\line_clipTable";
            string erase_table = $@"{def_gdb}\line_eraseTable";

            Arcpy.Erase(df, line, erase_area);
            // 裁剪汇总【line、erase_area】
            Arcpy.Clip(sd_clip, line, clip_fc);
            Arcpy.Statistics(clip_fc, clip_table, "shape_area SUM", gnField);
            Arcpy.CalculateField(clip_table, "SUM_shape_area", $"!SUM_shape_area!/10000*15/10000");
            Dictionary<string, double> dic_clip = GisTool.GetDictFromPathDouble(clip_table, gnField, "SUM_shape_area");
            ExcelTool.AttributeMapperDouble(excel_path, 0, col, dic_clip);

            Arcpy.Clip(sd_clip, erase_area, erase_fc);
            Arcpy.Statistics(erase_fc, erase_table, "shape_area SUM", gnField);
            Arcpy.CalculateField(erase_table, "SUM_shape_area", $"!SUM_shape_area!/10000*15/10000");
            Dictionary<string, double> dic_erase = GisTool.GetDictFromPathDouble(erase_table, gnField, "SUM_shape_area");
            ExcelTool.AttributeMapperDouble(excel_path, 0, col + 1, dic_erase);

        }

        private List<string> CheckData(List<string> lines, string sd)
        {
            List<string> result = new List<string>();


            // 检查DLBM的字段值
            string result_value = CheckTool.CheckFieldValue(sd, "DLBM", GlobalData.dic_sdAll.Keys.ToList());
            if (result_value != "")
            {
                result.Add(result_value);
            }

            // 检查是否正常提取Excel
            string result_excel = CheckTool.CheckExcelPick();
            if (result_excel != "")
            {
                result.Add(result_excel);
            }

            return result;
        }
    }
}
