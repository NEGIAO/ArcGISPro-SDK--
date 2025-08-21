using ArcGIS.Core.Data;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.SS.Formula.Functions;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using static Org.BouncyCastle.Math.Primes;

namespace CCTool.Scripts.DataPross.CAD
{
    /// <summary>
    /// Interaction logic for CADJZAnalysis.xaml
    /// </summary>
    public partial class CADJZAnalysis : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "CADJZAnalysis";

        public CADJZAnalysis()
        {
            InitializeComponent();

            // 初始化参数
            textCADPath.Text = BaseTool.ReadValueFromReg(toolSet, "cadPath");
            textExcelPath.Text = BaseTool.ReadValueFromReg(toolSet, "excelPath");
            textFeatureClassPath.Text = BaseTool.ReadValueFromReg(toolSet, "featureClassPath");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "CAD建筑分析";

        private void openCADButton_Click(object sender, RoutedEventArgs e)
        {
            textCADPath.Text = UITool.OpenDialogCAD();
        }

        private void saveFeatureClassButton_Click(object sender, RoutedEventArgs e)
        {
            textFeatureClassPath.Text = UITool.SaveDialogFeatureClass();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                var def_gdb = Project.Current.DefaultGeodatabasePath;
                // 获取指标
                string cadPath = textCADPath.Text;
                string excelPath = textExcelPath.Text;
                string featureClassPath = textFeatureClassPath.Text;
                bool isSimple = (bool)isSimpleCheck.IsChecked;

                // 判断参数是否选择完全
                if (cadPath == "" || featureClassPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "cadPath", cadPath);
                BaseTool.WriteValueToReg(toolSet, "excelPath", excelPath);
                BaseTool.WriteValueToReg(toolSet, "featureClassPath", featureClassPath);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 异步执行
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(featureClassPath);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10, "要素转面");
                    string cadPolyline = $@"{cadPath}\Polyline";
                    string polygonFeature = $@"{def_gdb}\polygonFeature";
                    Arcpy.FeatureToPolygon(cadPolyline, polygonFeature);

                    // 比较复杂的空间连接，需要文字连接
                    pw.AddMessageMiddle(10, "空间连接");
                    string cadAnnotation = $@"{cadPath}\Annotation";
                    string field = "TxtMemo";

                    GPExecuteToolFlags executeFlags = Arcpy.SetGPFlag(false);
                    string joinFeature = $@"{def_gdb}\joinFeature";

                    string exp = $"{field} '{field}' true true false 2048 Text 0 0,Join,#,{cadAnnotation},{field},0,2048";
                    var par = Geoprocessing.MakeValueArray(polygonFeature, cadAnnotation, joinFeature, "JOIN_ONE_TO_ONE", "KEEP_ALL", exp, "INTERSECT", "", "");
                    Geoprocessing.ExecuteToolAsync("analysis.SpatialJoin", par, null, null, null, executeFlags);

                    pw.AddMessageMiddle(20, "字段处理");
                    // 添加字段
                    string jg = "建筑结构";
                    string cs = "建筑层数";
                    string zl = "建筑质量";
                    string mj = "建筑面积";
                    Arcpy.AddField(joinFeature, jg, "TEXT");
                    Arcpy.AddField(joinFeature, cs, "SHORT");
                    Arcpy.AddField(joinFeature, zl, "TEXT");
                    Arcpy.AddField(joinFeature, mj, "DOUBLE");
                    // 计算字段
                    FieldCalTool.GetChinese(joinFeature, field, jg);
                    FieldCalTool.GetNumber(joinFeature, field, cs);
                    // 如果要简化结构
                    if (isSimple)
                    {
                        string block = "def ss(a):\r\n    b = a\r\n    if a is None or a == '':\r\n        b = '简'\r\n    elif a in ['砼','基','建','杦']:\r\n        b = '混'\r\n    elif a in ['破','棚','牲','厕','瓦']:\r\n        b = '简'\r\n    return b";
                        Arcpy.CalculateField(joinFeature, jg, $"ss(!{jg}!)", block);
                    }
                    // 简化层数
                    string block2 = "def ss(a):\r\n    if a is None:\r\n        return 1\r\n    else:\r\n        return a";
                    Arcpy.CalculateField(joinFeature, cs, $"ss(!{cs}!)", block2);

                    // 计算建筑质量
                    string block3 = "def ss(a,b):\r\n    if a == u'混' and b>2:\r\n        return '质量较好'\r\n    elif a == u'混' and b<=2:\r\n        return '质量一般'\r\n    else:\r\n        return '质量较差'";
                    Arcpy.CalculateField(joinFeature, zl, $"ss(!{jg}!,!{cs}!)", block3);

                    // 计算建筑面积
                    Arcpy.CalculateField(joinFeature, mj, $"!{cs}! * !shape.area!");

                    // 提纯
                    Arcpy.Select(joinFeature, featureClassPath, $"{jg} NOT IN ('天井', '水泥', '水泵', '水', '井')");
                    // 修复几何
                    Arcpy.RepairGeometry(featureClassPath);

                    pw.AddMessageMiddle(20, "统计面积");
                    // 复制excel
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.规划.建筑分析.xlsx", excelPath);

                    // 汇总
                    Dictionary<string, double> jgDict = ComboTool.StatisticsPlus(featureClassPath, jg, mj, "合计");
                    Dictionary<string, double> csDict = ComboTool.StatisticsPlus(featureClassPath, cs, mj, "合计");
                    Dictionary<string, double> zlDict = ComboTool.StatisticsPlus(featureClassPath, zl, mj, "合计");
                    // 写入excel
                    WriteExcel($@"{excelPath}\{jg}$", jgDict);
                    WriteExcel($@"{excelPath}\{cs}$", csDict);
                    WriteExcel($@"{excelPath}\{zl}$", zlDict);

                    pw.AddMessageMiddle(10, "清理数据并加载");
                    // 删除字段
                    List<string> keepFields = new List<string>() { jg, cs, zl, mj, field };
                    Arcpy.DeleteField(featureClassPath, keepFields, "KEEP_FIELDS");

                    // 删除中间数据
                    Arcpy.Delect(polygonFeature);

                    // 加载图层
                    MapCtlTool.AddLayerToMap(featureClassPath);
                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void WriteExcel(string excelPath, Dictionary<string, double> dict)
        {
            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            Cells cells = sheet.Cells;

            int index = 3;
            foreach (var pair in dict)
            {
                if (pair.Key != "合计")
                {
                    // 复制行，写入
                    cells.CopyRow(cells, 2, index);
                    cells[index, 1].Value = index - 2;
                    cells[index, 2].Value = pair.Key;
                    cells[index, 3].Value = pair.Value;
                    cells[index, 4].Value = pair.Value / dict["合计"] * 100;
                    index++;
                }
            }
            // 合计行
            cells.CopyRow(cells, 2, index);
            cells.Merge(index, 1, 1, 2);
            cells[index, 1].Value = "合计";
            cells[index, 3].Value = dict["合计"];
            cells[index, 4].Value = 100;

            cells.DeleteRow(2);

            // 保存
            wb.Save(excelFile);
            wb.Dispose();



        }



        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147139480";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string featureClassPath)
        {
            List<string> result = new List<string>();

            // 检查gdb要素路径是否合理
            string result_value = CheckTool.CheckGDBPath(featureClassPath);
            if (result_value != "")
            {
                result.Add(result_value);
            }

            string result_value2 = CheckTool.CheckGDBIsNumeric(featureClassPath);
            if (result_value2 != "")
            {
                result.Add(result_value);
            }
            return result;
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.SaveDialogExcel();
        }
    }
}
