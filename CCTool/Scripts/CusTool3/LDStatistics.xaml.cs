using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.GeoProcessing;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
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
using Row = ArcGIS.Core.Data.Row;
using Table = ArcGIS.Core.Data.Table;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for LDStatistics.xaml
    /// </summary>
    public partial class LDStatistics : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public LDStatistics()
        {
            InitializeComponent();
            textExcelPath.Text = $@"{Project.Current.HomeFolderPath}\导出图斑占三调用地统计表";
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "图斑占三调用地统计表(葛)";


        // 运行
        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                string defGDB = Project.Current.DefaultGeodatabasePath;
                string defFolder = Project.Current.HomeFolderPath;

                // 获取参数
                string sd = combox_sd.ComboxText();
                string xzq = combox_xzq.ComboxText();
                string txtPath = textTXTPath.Text;
                string excelFolder = textExcelPath.Text;


                // 判断参数是否选择完全
                if (sd == "" || xzq == "" || txtPath == "" || excelFolder == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 创建目标文件夹
                if (!Directory.Exists(excelFolder))
                {
                    Directory.CreateDirectory(excelFolder);
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("通过TXT文件创建图斑");
                    // 通过TXT文件创建图斑
                    string tb = CreateTB(txtPath);

                    pw.AddMessageMiddle(20, "相交并处理相关字段");
                    // 相交
                    List<string> fcs = new List<string>() { sd, tb };
                    string intersectResult = $@"{defGDB}\intersectResult";
                    Arcpy.Intersect(fcs, intersectResult);
                    // 排序地块名称字段
                    GisTool.AddField(intersectResult, "名称排序", FieldType.Integer);
                    string block = "import re\r\ndef ss(a):\r\n    va = re.findall(r'\\d', a)\r\n    if len(va) == 0:\r\n        result = ''\r\n    else:\r\n        result = ''\r\n        for i in range(0, len(va)):\r\n            result += va[i]\r\n    return result";
                    Arcpy.CalculateField(intersectResult, "名称排序", "ss(!地块编号!)", block);
                    // 排序
                    string sortResult = $@"{defGDB}\sortResult";
                    Arcpy.Sort(intersectResult, sortResult, "名称排序 ASCENDING;ZLDWMC ASCENDING;TBBH ASCENDING;DLMC ASCENDING;DLBM ASCENDING;QSDWMC ASCENDING", "UR");

                    pw.AddMessageMiddle(20, "写入【按地块】表");
                    string dkExcel = $@"{excelFolder}\按地块.xlsx";
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.图斑占林地统计表.按地块.xlsx", dkExcel);
                    // 写入【按地块】表
                    Write2DK(dkExcel, sortResult);

                    pw.AddMessageMiddle(20, "写入【按地块二级】表");
                    dkExcel = $@"{excelFolder}\按地块二级.xlsx";
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.图斑占林地统计表.按地块二级.xlsx", dkExcel);
                    // 写入【按地块二级】表
                    Write2DK2(dkExcel, sortResult, tb);

                    pw.AddMessageMiddle(20, "写入【按乡镇】表");
                    dkExcel = $@"{excelFolder}\按乡镇.xlsx";
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.图斑占林地统计表.按乡镇.xlsx", dkExcel);
                    // 写入【按乡镇】表
                    Write2XZ(dkExcel, sortResult, tb, xzq);

                });

                // 删除中间数据
                Arcpy.Delect($@"{defGDB}\intersectResult");
                Arcpy.Delect($@"{defGDB}\sortResult");

                Arcpy.Delect($@"{defGDB}\idenTarget");
                Arcpy.Delect($@"{defGDB}\out_table");

                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        // 写入【按乡镇】表
        private void Write2XZ(string dkExcel, string sortResult, string tb, string xzq)
        {
            string defGDB = Project.Current.DefaultGeodatabasePath;
            // 处理相交地块，标记乡镇名
            string idenTarget = $@"{defGDB}\idenTarget";
            Arcpy.Identity(sortResult, xzq, idenTarget);
            // 汇总
            string out_table = $@"{defGDB}\out_table";
            Arcpy.Statistics(idenTarget, out_table, "Shape_Area SUM", "名称排序;地块名称;XZQMC;QSDWMC;DLBM");
            // 添加一个标记字段
            Arcpy.AddField(out_table, "标记", "TEXT");
            Arcpy.CalculateField(out_table, "标记", "!地块名称!+!XZQMC!+!QSDWMC!");

            // 汇总各个地块的指标
            List<DKAtt> dKAtts = new List<DKAtt>();
            // 写入地块名称、标记
            Dictionary<string, string> bhmcDict = out_table.Get2FieldValueDic("标记", "地块名称");
            foreach (var bhmc in bhmcDict)
            {
                DKAtt dKAtt = new DKAtt();
                dKAtt.BJ = bhmc.Key;
                dKAtt.MC = bhmc.Value;
                dKAtts.Add(dKAtt);
            }
            // 写入乡镇名
            Dictionary<string, string> xzDic = out_table.Get2FieldValueDic("标记", "XZQMC");
            foreach (var dKAtt in dKAtts)
            {
                dKAtt.XZName = xzDic[dKAtt.BJ];
            }
            // 写入村庄名
            Dictionary<string, string> czDic = out_table.Get2FieldValueDic("标记", "QSDWMC");
            foreach (var dKAtt in dKAtts)
            {
                dKAtt.CZName = czDic[dKAtt.BJ];
            }
            // 写入地块总面积
            Dictionary<string, string> mjDic = out_table.Get2FieldValueDic("标记", "SUM_Shape_Area");
            foreach (var dKAtt in dKAtts)
            {
                dKAtt.MJ = double.Parse(mjDic[dKAtt.BJ]) / 10000;
            }
            // 写入各地类面积
            foreach (var dKAtt in dKAtts)
            {
                Dictionary<string, double> dlDic = ComboTool.StatisticsPlus(out_table, "DLBM", "SUM_Shape_Area", "合计", 10000, $"标记='{dKAtt.BJ}'");

                dKAtt.Dict = dlDic;
            }

            // 最后加一个合计项
            Dictionary<string, double> allDLDic = ComboTool.StatisticsPlus(out_table, "DLBM", "SUM_Shape_Area", "总计", 10000);
            DKAtt dKAttTotal = new DKAtt()
            {
                BH = "总计",
                MC = "总计",
                MJ = allDLDic["总计"],
                Dict = allDLDic,

                XZName = "总计",
                CZName = "总计",
                BJ = "总计",
            };
            dKAtts.Add(dKAttTotal);

            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(dkExcel);
            int sheetIndex = ExcelTool.GetSheetIndex(dkExcel);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取cell
            Cells cells = sheet.Cells;

            // 计数
            int index = 5;
            foreach (var dKAtt in dKAtts)
            {
                string mc = dKAtt.MC;
                string xz = dKAtt.XZName;
                string cz = dKAtt.CZName;

                double mj = dKAtt.MJ;
                Dictionary<string, double> dict = dKAtt.Dict;

                // 复制行
                cells.CopyRow(cells, 4, index);

                // 写入参数
                cells[index, 0].Value = mc;
                cells[index, 1].Value = xz;
                cells[index, 2].Value = cz;
                cells[index, 3].Value = mj;
                // 写入地类面积
                for (int i = 4; i < cells.MaxDataColumn + 1; i++)
                {
                    string dlbm = cells[3, i].StringValue;

                    if (dict.ContainsKey(dlbm))
                    {
                        cells[index, i].Value = dict[dlbm];
                    }
                }

                // 最后一行的话，合并格
                if (index == dKAtts.Count + 4)
                {
                    cells.Merge(index, 0, 1, 3);
                }
                index++;
            }

            // 删除示例行
            cells.DeleteRow(4);
            // 保存
            wb.Save(excelFile);
            wb.Dispose();

            // 删除空列
            ExcelTool.DeleteNullCol(dkExcel, dKAtts.Count + 3, 4);
        }

        // 写入【按地块二级】表
        private void Write2DK2(string dkExcel, string sortResult, string tb)
        {
            // 汇总各个地块的指标
            List<DKAtt> dKAtts = new List<DKAtt>();
            // 写入地块编号、地块名称
            Dictionary<string, string> bhmcDict = tb.Get2FieldValueDic("地块编号", "地块名称");
            foreach (var bhmc in bhmcDict)
            {
                DKAtt dKAtt = new DKAtt();
                dKAtt.BH = bhmc.Key;
                dKAtt.MC = bhmc.Value;
                dKAtts.Add(dKAtt);
            }
            // 写入地块总面积
            Dictionary<string, string> mjDic = tb.Get2FieldValueDic("地块编号", "shape_area");
            foreach (var dKAtt in dKAtts)
            {
                dKAtt.MJ = double.Parse(mjDic[dKAtt.BH]) / 10000;
            }
            // 写入各地类面积
            // 先把扩展地类处理一下
            string block = "def ss(a):\r\n    if len(a)>4:\r\n        return a[:4]\r\n    else:\r\n        return a";
            Arcpy.CalculateField(sortResult, "DLBM", "!DLBM!", block);

            foreach (var dKAtt in dKAtts)
            {
                Dictionary<string, double> dlDic = ComboTool.StatisticsPlus(sortResult, "DLBM", "shape_area", "合计", 10000, $"地块编号='{dKAtt.BH}'");

                dKAtt.Dict = dlDic;
            }

            // 最后加一个合计项
            Dictionary<string, double> allDLDic = ComboTool.StatisticsPlus(sortResult, "DLBM", "shape_area", "总计", 10000);
            DKAtt dKAttTotal = new DKAtt()
            {
                BH = "总计",
                MC = "总计",
                MJ = allDLDic["总计"],
                Dict = allDLDic
            };
            dKAtts.Add(dKAttTotal);

            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(dkExcel);
            int sheetIndex = ExcelTool.GetSheetIndex(dkExcel);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取cell
            Cells cells = sheet.Cells;

            // 计数
            int index = 5;
            foreach (var dKAtt in dKAtts)
            {
                string bh = dKAtt.BH;
                string mc = dKAtt.MC;
                double mj = dKAtt.MJ;
                Dictionary<string, double> dict = dKAtt.Dict;

                // 复制行
                cells.CopyRow(cells, 4, index);

                // 写入参数
                cells[index, 0].Value = bh;
                cells[index, 1].Value = mc;
                cells[index, 2].Value = mj;
                // 写入地类面积
                for (int i = 3; i < cells.MaxDataColumn + 1; i++)
                {
                    string dlbm = cells[3, i].StringValue;

                    if (dict.ContainsKey(dlbm))
                    {
                        cells[index, i].Value = dict[dlbm];
                    }
                }

                // 最后一行的话，合并格
                if (index == dKAtts.Count + 4)
                {
                    cells.Merge(index, 0, 1, 2);
                }
                index++;
            }

            // 删除示例行
            cells.DeleteRow(4);
            // 保存
            wb.Save(excelFile);
            wb.Dispose();

            // 删除空列
            ExcelTool.DeleteNullCol(dkExcel, dKAtts.Count + 3, 3);

        }


        // 写入【按地块】表
        private void Write2DK(string dkExcel, string sortResult)
        {
            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(dkExcel);
            int sheetIndex = ExcelTool.GetSheetIndex(dkExcel);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取cell
            Cells cells = sheet.Cells;

            Table table = sortResult.TargetTable();
            using RowCursor rowCursor = table.Search();
            // 计数
            int index = 1;

            while (rowCursor.MoveNext())
            {
                // 复制行
                if (index > 1)
                {
                    cells.CopyRow(cells, 1, index);
                }

                using Row row = rowCursor.Current;
                // 获取参数
                string DKMC = row["地块名称"]?.ToString();
                string TBBH = row["TBBH"]?.ToString();
                string DLBM = row["DLBM"]?.ToString();
                string DLMC = row["DLMC"]?.ToString();
                string QSDWMC = row["QSDWMC"]?.ToString();
                string ZLDWMC = row["ZLDWMC"]?.ToString();

                double MJ = double.Parse(row["shape_area"]?.ToString()) / 10000;

                // 写入参数
                cells[index, 0].Value = DKMC;
                cells[index, 1].Value = MJ;
                cells[index, 2].Value = ZLDWMC;
                cells[index, 3].Value = TBBH;
                cells[index, 4].Value = DLMC;
                cells[index, 5].Value = DLBM;
                cells[index, 6].Value = "地类图斑";
                cells[index, 7].Value = QSDWMC;
                index++;
            }

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // 通过TXT文件创建图斑
        public string CreateTB(string txtPath)
        {
            string gdb_def = Project.Current.DefaultGeodatabasePath;

            // 获取txt文件的文本内容
            string text = TxtTool.GetTXTContent(txtPath);
            // 文本中的【@】符号放前
            string updata_text = ChangeSymbol(text);

            // 提取属性描述
            Dictionary<string, string> dict = GetAtt(text);
            // 获取坐标系
            string spatial_reference = $"CGCS2000_3_Degree_GK_Zone_{dict["带号"]}";

            // 创建一个空要素
            Arcpy.CreateFeatureclass(gdb_def, "tem_fc", "POLYGON", spatial_reference);
            string targetFC = gdb_def + @"\tem_fc";

            // 新建字段
            GisTool.AddField(targetFC, "地块编号", FieldType.String);
            GisTool.AddField(targetFC, "地块名称", FieldType.String);

            // 打开数据库
            using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_def))))
            {
                // 创建要素并添加到要素类中
                using FeatureClass featureClass = gdb.OpenDataset<FeatureClass>("tem_fc");
                // 解析txt文件内容，创建面要素
                // 获取坐标点文本
                string[] fcs_text = updata_text.Split("@");
                // 去除第一部分非坐标文本
                List<string> fcs_text2List = new List<string>(fcs_text);
                fcs_text2List.RemoveAt(0);

                // 一个文件可能有多要素
                foreach (var txt in fcs_text2List)
                {
                    // 获取要素的部件数
                    int parts = GetCount(txt);

                    // 构建坐标点集合
                    var vertices_list = new List<List<Coordinate2D>>();
                    for (int i = 0; i < parts; i++)
                    {
                        var vertices = new List<Coordinate2D>();
                        vertices_list.Add(vertices);
                    }

                    // 编号、名称
                    string dkbh = "";
                    string dkmc = "";
                    // 根据换行符分解坐标点文本
                    string[] list_point = txt.Split("\n");

                    // 监视部件号变化
                    string partID = "-1";
                    int pID = -1;

                    foreach (var point in list_point)
                    {
                        if (TxtTool.StringInCount(point, ",") == 8)     // 名称、地块编号、功能文本
                        {
                            dkbh = point.Split(",")[2];
                            dkmc = point.Split(",")[3];
                        }
                        else if (TxtTool.StringInCount(point, ",") == 3)           // 点坐标文本
                        {
                            string fid = point.Split(",")[1].Replace(" ", "");        // 图斑部件号
                            if (fid != partID)
                            {
                                pID += 1;
                                partID = fid;
                            }
                            double lat = double.Parse(point.Split(",")[3].Replace(" ", ""));         // 经度
                            double lng = double.Parse(point.Split(",")[2].Replace(" ", ""));         // 纬度

                            vertices_list[pID].Add(new Coordinate2D(lat, lng));    // 加入坐标点集合
                        }
                        else     // 跳过无坐标部份的文本
                        {
                            continue;
                        }
                    }

                    /// 构建面要素
                    // 创建编辑操作对象
                    EditOperation editOperation = new EditOperation();
                    editOperation.Callback(context =>
                    {
                        // 获取要素定义
                        FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();
                        // 创建RowBuffer
                        using RowBuffer rowBuffer = featureClass.CreateRowBuffer();

                        // 写入字段值
                        rowBuffer["地块编号"] = dkbh;
                        rowBuffer["地块名称"] = dkmc;

                        PolygonBuilderEx pb = new PolygonBuilderEx(vertices_list[0]);
                        // 如果有空洞，则添加内部Polygon
                        if (vertices_list.Count > 1)
                        {
                            for (int i = 0; i < vertices_list.Count - 1; i++)
                            {
                                pb.AddPart(vertices_list[i + 1]);
                            }
                        }
                        // 给新添加的行设置形状
                        rowBuffer[featureClassDefinition.GetShapeField()] = pb.ToGeometry();

                        // 在表中创建新行
                        using Feature feature = featureClass.CreateRow(rowBuffer);
                        context.Invalidate(feature);      // 标记行为无效状态
                    }, featureClass);

                    // 执行编辑操作
                    editOperation.Execute();
                }
            }
            // 保存编辑
            Project.Current.SaveEditsAsync();
            // 修复几何
            Arcpy.RepairGeometry(targetFC);

            return targetFC;
        }



        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/144261005?spm=1001.2014.3001.5502";
            UITool.Link2Web(url);
        }


        private void combox_sd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_sd);
        }

        private void combox_xzq_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_xzq);
        }

        private void openTXTButton_Click(object sender, RoutedEventArgs e)
        {
            textTXTPath.Text = UITool.OpenDialogTXT();
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.OpenDialogFolder();
        }

        // 文本中的【@】符号放前
        public static string ChangeSymbol(string text)
        {
            string[] lins = text.Split('\n');
            string updata_lins = "";
            foreach (string line in lins)
            {

                if (line.Contains("@"))
                {
                    string newline = line.Replace("@", "");
                    newline = "@" + newline;
                    updata_lins += newline + "\n";
                }
                else
                {
                    updata_lins += line + "\n";
                }
            }
            return updata_lins;
        }

        // 获取指标
        public static Dictionary<string, string> GetAtt(string text)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            string new_text = text[text.IndexOf("[属性描述]")..text.IndexOf("[地块坐标]")];
            string[] lines = new_text.Split("\n");
            foreach (string line in lines)
            {
                if (line.Contains('='))
                {
                    string before = line[..line.IndexOf("=")];
                    string after = line[(line.IndexOf("=") + 1)..].Replace("\r", "");
                    dict.Add(before, after);
                }
            }
            return dict;
        }

        // 获取要素的部件数
        public static int GetCount(string lines)
        {
            List<string> indexs = new List<string>();

            // 根据换行符分解坐标点文本
            string[] list_point = lines.Split("\n");

            foreach (var point in list_point)
            {
                if (TxtTool.StringInCount(point, ",") == 3)          // 点坐标文本
                {
                    // 判断是否带空洞
                    string fid = point.Split(",")[1];        // 图斑部件号
                    if (!indexs.Contains(fid))
                    {
                        indexs.Add(fid);
                    }
                }
                else    // 路过非点坐标文本
                {
                    continue;
                }
            }

            return indexs.Count;
        }

    }
}

// 地块属性
public class DKAtt
{
    public string MC { get; set; }
    public string BH { get; set; }
    public double MJ { get; set; }
    public Dictionary<string, double> Dict { get; set; }

    public string XZName { get; set; }
    public string CZName { get; set; }

    public string BJ { get; set; }

}
