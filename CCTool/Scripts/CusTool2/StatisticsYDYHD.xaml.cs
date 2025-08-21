using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Windows;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using static ArcGIS.Desktop.Internal.Core.PortalTrafficDataService.ServiceErrorResponse;
using Brushes = System.Windows.Media.Brushes;
using Row = ArcGIS.Core.Data.Row;

namespace CCTool.Scripts.CusTool2
{
    /// <summary>
    /// Interaction logic for StatisticsYDYHD.xaml
    /// </summary>
    public partial class StatisticsYDYHD : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public StatisticsYDYHD()
        {
            InitializeComponent();
            Init();       // 初始化
        }

        // 初始化
        private void Init()
        {
            // combox_model框中添加3种转换模式，默认【中类】
            combox_model.Items.Add("大类");
            combox_model.Items.Add("中类");
            combox_model.Items.Add("小类");
            combox_model.SelectedIndex = 1;

            combox_unit.Items.Add("平方米");
            combox_unit.Items.Add("公顷");
            combox_unit.Items.Add("平方公里");
            combox_unit.Items.Add("亩");
            combox_unit.SelectedIndex = 1;

            combox_digit.Items.Add("1");
            combox_digit.Items.Add("2");
            combox_digit.Items.Add("3");
            combox_digit.Items.Add("4");
            combox_digit.Items.Add("5");
            combox_digit.Items.Add("6");
            combox_digit.SelectedIndex = 1;

            textExcelPath.Text = BaseTool.ReadValueFromReg("YDYHD", "excelPath");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "用地用海指标汇总(栋)";

        // 点击打开按钮，选择输出的Excel文件位置
        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            // 打开Excel文件
            string path = UITool.SaveDialogExcel();
            // 将Excel文件的路径置入【textExcelPath】
            textExcelPath.Text = path;

            BaseTool.WriteValueToReg("YDYHD", "excelPath", path);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        // 运行
        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                var init_gdb = Project.Current.DefaultGeodatabasePath;
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string field_mc = combox_mc.ComboxText();
                string excel_path = textExcelPath.Text;

                string model = combox_model.Text;

                string unit = combox_unit.Text;
                int digit = int.Parse(combox_digit.Text);

                string kfbj = combox_kfbj.ComboxText();
                string qs = combox_qs.ComboxText();

                bool isKFBJ = (bool)cb_kfbj.IsChecked;
                bool isQS = (bool)cb_qs.IsChecked;

                // 判断参数是否选择完全
                if (fc_path == "" || field_mc == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }
                if (isKFBJ == true && kfbj == "")
                {
                    MessageBox.Show("开发边界已启用，请选择相应图层！");
                    return;
                }
                if (isQS == true && qs == "")
                {
                    MessageBox.Show("权属已启用，请选择相应图层！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                // 单位系数设置
                double unit_xs = unit switch
                {
                    "平方米" => 1,
                    "公顷" => 10000,
                    "平方公里" => 1000000,
                    "亩" => 666.66667,
                    _ => 1,
                };

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(fc_path, field_mc, qs);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(0, err, Brushes.Red);
                        }
                        return;
                    }

                    // 选择合适的模板
                    string excelName = "";
                    if (isKFBJ)
                    {
                        if (model == "大类")
                        {
                            excelName = "【新模板】用地用海_大类_开发边界";
                        }
                        else if (model == "中类")
                        {
                            excelName = "【新模板】用地用海_中类_开发边界";
                        }
                        else if (model == "小类")
                        {
                            excelName = "【新模板】用地用海_小类_开发边界";
                        }
                    }
                    else
                    {
                        if (model == "大类")
                        {
                            excelName = "【新模板】用地用海_大类_无开发边界";
                        }
                        else if (model == "中类")
                        {
                            excelName = "【新模板】用地用海_中类_无开发边界";
                        }
                        else if (model == "小类")
                        {
                            excelName = "【新模板】用地用海_小类_无开发边界";
                        }
                    }
                    
                    // 复制嵌入资源中的Excel文件
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.{excelName}.xlsx", excel_path);
                    string excel_sheet = excel_path + @"\用地用海$";
                    // 设置小数位数
                    ExcelTool.SetDigit(excel_sheet, new List<int>() { 4 }, 4, digit);

                    // 设置表格位置参数
                    int mapCol = 7;       // 参照列
                    int mjCol = 0;        // 面积列
                    if (model == "大类") { mjCol = 3; }
                    else if (model == "中类") { mjCol = 4; }
                    else if (model == "小类") { mjCol = 5; }
                    int blCol = mjCol + 1;        // 占比列
                    int startRow = 4;       // 起始行
                    int endRow = 0;       // 末尾行
                    if (model == "大类") { endRow = 28; }
                    else if (model == "中类") { endRow = 139; }
                    else if (model == "小类") { endRow = 179; }
                    int startRow2 = endRow+ 1;
                    int endRow2 = endRow-startRow+startRow2;

                    // 有开发边界的情况
                    if (isKFBJ)
                    {
                        pw.AddMessageMiddle(20, "标识开发边界和权属");
                        string iden_kfbj = $@"{init_gdb}\iden_kfbj";
                        // 字段预处理
                        string fidField = "";
                        if (kfbj.Contains("\\"))
                        {
                            int index = kfbj.LastIndexOf("\\");
                            fidField = $"FID_{kfbj[(index + 1)..]}";
                        }
                        else
                        {
                            fidField = $"FID_{kfbj}";
                        }

                        if (GisTool.IsHaveFieldInTarget(kfbj, fidField))
                        {
                            GisTool.DeleteField(kfbj, fidField);
                        }
                        // 标识开发边界和权属图层
                        Arcpy.Identity(qs, kfbj, iden_kfbj);
                        // 添加并计算一个标记字段【BZ】
                        if (!GisTool.IsHaveFieldInTarget(iden_kfbj, "BZ"))
                        {
                            Arcpy.AddField(iden_kfbj, "BZ", "TEXT");
                        }
                        // 计算标记字段
                        string block = "def ss(a):\r\n    if a == -1:\r\n        return \"开发边界外\"\r\n    else:\r\n        return \"开发边界内\"";
                        Arcpy.CalculateField(iden_kfbj, "BZ", $"ss(!{fidField}!)", block);

                        pw.AddMessageMiddle(20, "汇总统计");
                        // 交集制表
                        List<string> fields = new List<string>() { "BZ", "XZQMC" };
                        string outTable = $@"{init_gdb}\outTable";
                        Arcpy.TabulateIntersection(iden_kfbj, fields, fc_path, outTable, field_mc);
                        // 添加编码字段
                        Arcpy.AddField(outTable, "一级编码", "TEXT");
                        Arcpy.AddField(outTable, "二级编码", "TEXT");
                        Arcpy.AddField(outTable, "三级编码", "TEXT");

                        // 名称编号键值对调
                        Dictionary<string, string> dic = GlobalData.dic_ydyh_new.ToDictionary(k => k.Value, p => p.Key);
                        // 计算字段
                        Table table = outTable.TargetTable();
                        using (RowCursor rowCursor = table.Search())
                        {
                            while (rowCursor.MoveNext())
                            {
                                using Row row = rowCursor.Current;
                                var mc_ob = row[field_mc];
                                if (mc_ob != null)
                                {
                                    string mc = mc_ob.ToString();
                                    string bm = dic[mc];
                                    string bm1 = bm[..2];
                                    string bm2 = bm.Length > 2 ? bm[..4] : "";
                                    string bm3 = bm.Length > 4 ? bm: "";
                                    row["一级编码"] = bm1;
                                    row["二级编码"] = bm2;
                                    row["三级编码"] = bm3;
                                }
                                row.Store();
                            }
                        }

                        pw.AddMessageMiddle(20, "写入Excel");
                        List<string> bmFields = new List<string>() { "一级编码", "二级编码",  "三级编码"};
                        // 分权属
                        if (isQS)
                        {
                            List<string> qsList = outTable.GetFieldValues("XZQMC");
                            // 分村计写入Excel
                            foreach (string qsmc in qsList)
                            {
                                if (qsmc is null || qsmc == "")
                                {
                                    continue;
                                }

                                pw.AddMessageMiddle(0, $"      {qsmc}", Brushes.Gray);
                                // 复制表格
                                ExcelTool.CopySheet(excel_path, "用地用海", qsmc);

                                // 打开工作薄
                                Workbook wb = ExcelTool.OpenWorkbook(excel_path);
                                // 打开工作表
                                Worksheet sheet = wb.Worksheets[qsmc];

                                // 开发边界外
                                Dictionary<string, double> dic_out = ComboTool.StatisticsPlus(outTable, bmFields, "shape_area", "合计", unit_xs, $"BZ = '开发边界外' AND XZQMC = '{qsmc}'");
                                double total_mj_out = dic_out.ContainsKey("合计") ? dic_out["合计"] : 0;
                                for (int i = startRow; i <= endRow; i++)
                                {
                                    //  获取目标cell
                                    Cell inCell = sheet.Cells[i, mapCol];
                                    Cell mapCell = sheet.Cells[i, mjCol];
                                    // 属性映射
                                    if (inCell is not null && dic_out.ContainsKey(inCell.StringValue))
                                    {
                                        double mapValue = dic_out[inCell.StringValue];
                                        mapCell.Value = mapValue;   // 赋值
                                        sheet.Cells[i, blCol].Value = Math.Round(mapValue / total_mj_out * 100, 2);
                                    }
                                }

                                // 开发边界内
                                Dictionary<string, double> dic_in = ComboTool.StatisticsPlus(outTable, bmFields, "shape_area", "合计", unit_xs, $"BZ = '开发边界内' AND XZQMC = '{qsmc}'");
                                double total_mj_in = dic_in.ContainsKey("合计") ? dic_in["合计"] : 0;
                                for (int i = startRow2; i <= endRow2; i++)
                                {
                                    //  获取目标cell
                                    Cell inCell = sheet.Cells[i, mapCol];
                                    Cell mapCell = sheet.Cells[i, mjCol];
                                    // 属性映射
                                    if (inCell is not null && dic_in.ContainsKey(inCell.StringValue))
                                    {
                                        double mapValue = dic_in[inCell.StringValue];
                                        mapCell.Value = mapValue;   // 赋值
                                        sheet.Cells[i, blCol].Value = Math.Round(mapValue / total_mj_in * 100, 2);
                                    }
                                }

                                // 保存
                                wb.Save(excel_path);
                                wb.Dispose();

                                string excelFullPath = @$"{excel_path}\{qsmc}$";
                                // 删除0值行
                                ExcelTool.DeleteNullRow(excelFullPath, mjCol, startRow);

                                // 删除指定列
                                ExcelTool.DeleteCol(excelFullPath, mapCol);
                                // 改Excel中的单位
                                ExcelTool.WriteCell(excelFullPath, 2, mjCol, $"用地面积({unit})");
                            }

                            ExcelTool.DeleteSheet(excel_path, "用地用海");
                        }
                        // 不分权属
                        else
                        {
                            // 打开工作薄
                            Workbook wb = ExcelTool.OpenWorkbook(excel_path);
                            // 打开工作表
                            Worksheet sheet = wb.Worksheets["用地用海"];

                            // 开发边界外
                            Dictionary<string, double> dic_out = ComboTool.StatisticsPlus(outTable, bmFields, "shape_area", "合计", unit_xs, $"BZ = '开发边界外'");
                            double total_mj_out = dic_out.ContainsKey("合计") ? dic_out["合计"] : 0;
                            for (int i = startRow; i <= endRow; i++)
                            {
                                //  获取目标cell
                                Cell inCell = sheet.Cells[i, mapCol];
                                Cell mapCell = sheet.Cells[i, mjCol];
                                // 属性映射
                                if (inCell is not null && dic_out.ContainsKey(inCell.StringValue))
                                {
                                    double mapValue = dic_out[inCell.StringValue];
                                    mapCell.Value = mapValue;   // 赋值
                                    sheet.Cells[i, blCol].Value = Math.Round(mapValue / total_mj_out * 100, 2);
                                }
                            }

                            // 开发边界内
                            Dictionary<string, double> dic_in = ComboTool.StatisticsPlus(outTable, bmFields, "shape_area", "合计", unit_xs, $"BZ = '开发边界内'");
                            double total_mj_in = dic_in.ContainsKey("合计") ? dic_in["合计"] : 0;
                            for (int i = startRow2; i <= endRow2; i++)
                            {
                                //  获取目标cell
                                Cell inCell = sheet.Cells[i, mapCol];
                                Cell mapCell = sheet.Cells[i, mjCol];
                                // 属性映射
                                if (inCell is not null && dic_in.ContainsKey(inCell.StringValue))
                                {
                                    double mapValue = dic_in[inCell.StringValue];
                                    mapCell.Value = mapValue;   // 赋值
                                    sheet.Cells[i, blCol].Value = Math.Round(mapValue / total_mj_in * 100, 2);
                                }
                            }

                            // 保存
                            wb.Save(excel_path);
                            wb.Dispose();

                            string excelFullPath = @$"{excel_path}\用地用海$";
                            // 删除0值行
                            ExcelTool.DeleteNullRow(excelFullPath, mjCol, startRow);

                            // 删除指定列
                            ExcelTool.DeleteCol(excelFullPath, mapCol);
                            // 改Excel中的单位
                            ExcelTool.WriteCell(excelFullPath, 2, mjCol, $"用地面积({unit})");
                        }
                    }
                    // 没有开发边界的情况
                    else
                    {
                        pw.AddMessageMiddle(20, "汇总统计");
                        // 交集制表
                        List<string> fields = new List<string>() { "XZQMC" };
                        string outTable = $@"{init_gdb}\outTable";
                        Arcpy.TabulateIntersection(qs, fields, fc_path, outTable, field_mc);
                        // 添加编码字段
                        Arcpy.AddField(outTable, "一级编码", "TEXT");
                        Arcpy.AddField(outTable, "二级编码", "TEXT");
                        Arcpy.AddField(outTable, "三级编码", "TEXT");

                        // 名称编号键值对调
                        Dictionary<string, string> dic = GlobalData.dic_ydyh_new.ToDictionary(k => k.Value, p => p.Key);
                        // 计算字段
                        Table table = outTable.TargetTable();
                        using (RowCursor rowCursor = table.Search())
                        {
                            while (rowCursor.MoveNext())
                            {
                                using Row row = rowCursor.Current;
                                var mc_ob = row[field_mc];
                                if (mc_ob != null)
                                {
                                    string mc = mc_ob.ToString();
                                    string bm = dic[mc];
                                    string bm1 = bm[..2];
                                    string bm2 = bm.Length > 2 ? bm[..4] : "";
                                    string bm3 = bm.Length > 4 ? bm : "";
                                    row["一级编码"] = bm1;
                                    row["二级编码"] = bm2;
                                    row["三级编码"] = bm3;
                                }
                                row.Store();
                            }
                        }

                        pw.AddMessageMiddle(20, "写入Excel");
                        List<string> bmFields = new List<string>() { "一级编码", "二级编码", "三级编码" };
                        // 分权属
                        if (isQS)
                        {
                            List<string> qsList = outTable.GetFieldValues("XZQMC");
                            // 分村计写入Excel
                            foreach (string qsmc in qsList)
                            {
                                if (qsmc is null || qsmc == "")
                                {
                                    continue;
                                }

                                pw.AddMessageMiddle(0, $"      {qsmc}", Brushes.Gray);
                                // 复制表格
                                ExcelTool.CopySheet(excel_path, "用地用海", qsmc);

                                // 打开工作薄
                                Workbook wb = ExcelTool.OpenWorkbook(excel_path);
                                // 打开工作表
                                Worksheet sheet = wb.Worksheets[qsmc];

                                // 开发边界
                                Dictionary<string, double> dic_out = ComboTool.StatisticsPlus(outTable, bmFields, "shape_area", "合计", unit_xs, $"XZQMC = '{qsmc}'");
                                double total_mj_out = dic_out.ContainsKey("合计") ? dic_out["合计"] : 0;
                                for (int i = startRow; i <= endRow; i++)
                                {
                                    //  获取目标cell
                                    Cell inCell = sheet.Cells[i, mapCol];
                                    Cell mapCell = sheet.Cells[i, mjCol];
                                    // 属性映射
                                    if (inCell is not null && dic_out.ContainsKey(inCell.StringValue))
                                    {
                                        double mapValue = dic_out[inCell.StringValue];
                                        mapCell.Value = mapValue;   // 赋值
                                        sheet.Cells[i, blCol].Value = Math.Round(mapValue / total_mj_out * 100, 2);
                                    }
                                }

                                // 保存
                                wb.Save(excel_path);
                                wb.Dispose();

                                string excelFullPath = @$"{excel_path}\{qsmc}$";
                                // 删除0值行
                                ExcelTool.DeleteNullRow(excelFullPath, mjCol, startRow);

                                // 删除指定列
                                ExcelTool.DeleteCol(excelFullPath, mapCol);
                                // 改Excel中的单位
                                ExcelTool.WriteCell(excelFullPath, 2, mjCol, $"用地面积({unit})");
                            }

                            ExcelTool.DeleteSheet(excel_path, "用地用海");
                        }
                        // 不分权属
                        else
                        {
                            // 打开工作薄
                            Workbook wb = ExcelTool.OpenWorkbook(excel_path);
                            // 打开工作表
                            Worksheet sheet = wb.Worksheets["用地用海"];

                            // 开发边界
                            Dictionary<string, double> dic_out = ComboTool.StatisticsPlus(outTable, bmFields, "shape_area", "合计", unit_xs);
                            double total_mj_out = dic_out.ContainsKey("合计") ? dic_out["合计"] : 0;
                            for (int i = startRow; i <= endRow; i++)
                            {
                                //  获取目标cell
                                Cell inCell = sheet.Cells[i, mapCol];
                                Cell mapCell = sheet.Cells[i, mjCol];
                                // 属性映射
                                if (inCell is not null && dic_out.ContainsKey(inCell.StringValue))
                                {
                                    double mapValue = dic_out[inCell.StringValue];
                                    mapCell.Value = mapValue;   // 赋值
                                    sheet.Cells[i, blCol].Value = Math.Round(mapValue / total_mj_out * 100, 2);
                                }
                            }

                            // 保存
                            wb.Save(excel_path);
                            wb.Dispose();

                            string excelFullPath = @$"{excel_path}\用地用海$";
                            // 删除0值行
                            ExcelTool.DeleteNullRow(excelFullPath, mjCol, startRow);

                            // 删除指定列
                            ExcelTool.DeleteCol(excelFullPath, mapCol);
                            // 改Excel中的单位
                            ExcelTool.WriteCell(excelFullPath, 2, mjCol, $"用地面积({unit})");
                        }
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

        private void combox_fc_Closed(object sender, EventArgs e)
        {
            try
            {
                // 填写输出路径
                textExcelPath.Text = Project.Current.HomeFolderPath + @"\用地用海指标汇总表.xlsx";
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }


        private List<string> CheckData(string in_data, string in_field, string qs)
        {
            List<string> result = new List<string>();


            if (in_data != "" && in_field != "")
            {
                // 检查字段值是否符合要求
                string result_value = CheckTool.CheckFieldValue(in_data, in_field, GlobalData.dic_ydyh_new.Values.ToList());
                if (result_value != "")
                {
                    result.Add(result_value);
                }

            }

            if (qs != "")
            {
                // 检查字段
                string result_value = CheckTool.IsHaveFieldInLayer(qs, "XZQMC");
                if (result_value != "")
                {
                    result.Add(result_value);
                }
            }

            return result;
        }

        private void combox_kfbj_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_kfbj);
        }

        private void combox_qs_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_qs);
        }

        private void combox_mc_DropDown(object sender, EventArgs e)
        {
            string fc_path = combox_fc.ComboxText();
            UITool.AddTextFieldsToComboxPlus(fc_path, combox_mc);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/143853272";
            UITool.Link2Web(url);
        }
    }
}
