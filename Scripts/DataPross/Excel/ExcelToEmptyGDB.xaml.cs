using ArcGIS.Core.Data;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Windows;
using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Tsp;
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
using System.Xml;

namespace CCTool.Scripts.UI.ProWindow
{
    /// <summary>
    /// Interaction logic for ExcelToEmptyGDB.xaml
    /// </summary>
    public partial class ExcelToEmptyGDB : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public ExcelToEmptyGDB()
        {
            InitializeComponent();
            Init();
        }

        // 初始化
        private void Init()
        {
            // combox_sr框中添加几种预制坐标系
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_30");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_31");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_32");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_33");

            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_34");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_35");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_36");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_37");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_38");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_39");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_40");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_41");

            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_90E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_93E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_96E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_99E");

            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_102E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_105E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_108E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_111E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_114E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_117E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_120E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_123E");
            combox_sr.SelectedIndex = 5;

            combox_sr_ele.Items.Add("");
            combox_sr_ele.Items.Add("Yellow_Sea_1985");
            combox_sr_ele.SelectedIndex = 0;

            combox_XYTolerance.Items.Add("0.001");
            combox_XYTolerance.Items.Add("0.0001");
            combox_XYTolerance.SelectedIndex = 0;

            combox_XYResolution.Items.Add("0.0001");
            combox_XYResolution.Items.Add("0.00005");
            combox_XYResolution.SelectedIndex = 0;
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "属性结构描述表转空库(批量)";

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.OpenDialogExcel();

            // 填写输出路径
            string text = textExcelPath.Text;
            if (text == "")
            {
                return;
            }

            textGDBPath.Text = text[..text.LastIndexOf(@"\")] + @"\输出空库";
        }

        private List<string> CheckData(string text)
        {
            List<string> result = new List<string>();

            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(text);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);

            // 针对每个工作表进行处理
            for (int i = 0; i < wb.Worksheets.Count; i++)
            {
                // 获取工作表
                Worksheet sheet = wb.Worksheets[i];

                // 解析表名，获取要素名、要素别名
                string name_feature = sheet.Cells[0, 1].StringValue;
                string name_alias = sheet.Cells[1, 1].StringValue;
                string fc_type = sheet.Cells[2, 1].StringValue;
                //  空值检查
                if (name_feature == "")
                {
                    result.Add($"【{name_alias}】要素名称为空");
                }
                if (name_alias == "")
                {
                    result.Add($"【{name_alias}】要素别名为空");
                }
                if (fc_type == "")
                {
                    result.Add($"【{name_alias}】要素类型为空");
                }
                List<string> tps = new List<string>() { "面", "线", "点", "表" };
                if (!tps.Contains(fc_type))
                {
                    result.Add($"【{name_alias}】要素类型填写不规范");
                }

                // 获取总行数
                int row_count = sheet.Cells.MaxDataRow;
                // 检查字段属性
                for (int j = 4; j <= row_count; j++)
                {
                    // 获取字段属性
                    string mc = sheet.Cells[j, 1].StringValue.Replace(" ", "");
                    string dm = sheet.Cells[j, 2].StringValue.Replace(" ", "");
                    string tp = sheet.Cells[j, 3].StringValue.Replace(" ", "");
                    string length = sheet.Cells[j, 4].StringValue.Replace(" ", "");
                    // 名称
                    if (mc == "")
                    {
                        result.Add($"【{name_alias}】第{j - 3}个字段名称为空");
                    }
                    // 代码
                    if (dm == "")
                    {
                        result.Add($"【{name_alias}】第{j - 3}个字段代码为空");
                    }
                    // 类型
                    List<string> fdTps = new List<string>()
                    {
                        "Char","VarChar","Text","String","Int","Long","Float","Double","Date",
                    };
                    if (!fdTps.Contains(tp))
                    {
                        result.Add($"【{name_alias}】第{j - 3}个字段类型不规范");
                    }
                    // 长度
                    if (length != "")
                    {
                        try
                        {
                            int len = int.Parse(length);
                        }
                        catch (Exception)
                        {
                            result.Add($"【{name_alias}】第{j - 3}个字段长度填写有误");
                            continue;
                        }
                    }

                }
            }

            return result;
        }


        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textGDBPath.Text = UITool.OpenDialogFolder();

        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string excel_path = textExcelPath.Text;
                string gdb_path = textGDBPath.Text;
                string spatial_reference = combox_sr.Text;
                string sr_elevation = combox_sr_ele.Text;
                bool isToDouble = (bool)checkDouble.IsChecked;

                double xyTolerance = double.Parse(combox_XYTolerance.Text);
                double xyResolution = double.Parse(combox_XYResolution.Text);

                // 提取Excel文件名【即数据库名】
                string name_excel = excel_path[(excel_path.LastIndexOf(@"\") + 1)..excel_path.LastIndexOf(@".")];

                // 判断参数是否选择完全
                if (excel_path == "" || gdb_path == "" || spatial_reference == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                // 异步执行
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(excel_path);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10, "创建空GDB数据库");
                    // 创建一个空的GDB数据库
                    Arcpy.CreateFileGDB(gdb_path, name_excel);

                    // 获取工作薄、工作表
                    string excelFile = ExcelTool.GetPath(excel_path);
                    // 打开工作薄
                    Workbook wb = ExcelTool.OpenWorkbook(excelFile);

                    // 针对每个工作表进行处理
                    for (int i = 0; i < wb.Worksheets.Count; i++)
                    {
                        // 获取工作表
                        Worksheet sheet = wb.Worksheets[i];

                        // 解析表名，获取要素名、要素别名
                        string name_feature = sheet.Cells[0, 1].StringValue.Replace(" ","");
                        string name_alias = sheet.Cells[1, 1].StringValue.Replace(" ", "");
                        string fc_type = sheet.Cells[2, 1].StringValue.Replace(" ", "");

                        pw.AddMessageMiddle(10, "转表：" + name_feature);

                        if (fc_type != "表")           //  要素类的情况
                        {
                            string sr = spatial_reference;
                            if (sr_elevation != "")   // 如果有垂直坐标系
                            {
                                sr = GlobalData.cdDic[spatial_reference] + "," + GlobalData.cdDic[sr_elevation];
                            }
                            string fc_type_final = GetFeatureClassType(fc_type);

                            // 创建空要素
                            Arcpy.CreateFeatureclass(gdb_path + @"\" + name_excel + @".gdb", name_feature, fc_type_final, sr, name_alias, xyTolerance, xyResolution);
                        }
                        else    //  表的情况
                        {
                            // 创建表
                            Arcpy.CreateTable(gdb_path + @"\" + name_excel + @".gdb", name_feature, name_alias);
                        }

                        // 获取总行数
                        int row_count = sheet.Cells.MaxDataRow;
                        // 创建字段
                        for (int j = 4; j <= row_count; j++)
                        {
                            // 获取字段属性
                            string mc = sheet.Cells[j, 1].StringValue;
                            string dm = sheet.Cells[j, 2].StringValue;
                            string length = sheet.Cells[j, 4].StringValue;
                            int length_value;
                            // 如果长度为空，则设为默认的255
                            if (length == "")
                            {
                                length_value = 255;
                            }
                            else
                            {
                                length_value = int.Parse(length);
                            }
                            if (mc != "" && dm != "")
                            {
                                string aliasName = mc.ToString().Replace("\n", "");               // 字段别名
                                string fieldName = dm.ToString().Replace("\n", "");              // 字段名
                                string field_type = GetFeildType(sheet.Cells[j, 3].StringValue, isToDouble);          // 字段类型
                                // 创建字段
                                string tartgetFC = gdb_path + @"\" + name_excel + @".gdb\" + name_feature;
                                Arcpy.AddField(tartgetFC, fieldName, field_type, aliasName, length_value);

                            }
                        }
                    }
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

        // 要素类型转换
        public static string GetFeatureClassType(string type)
        {
            string fc_type = "";
            if (type == "点") { fc_type = "POINT"; }
            else if (type == "线") { fc_type = "POLYLINE"; }
            else if (type == "面") { fc_type = "POLYGON"; }

            return fc_type;
        }

        // 字段类型转换
        public static FieldType GetFeildType2(string type)
        {
            FieldType fd_type = FieldType.String;
            if (type == "Float") { fd_type = FieldType.Double; }
            else if (type == "Char" || type == "VarChar") { fd_type = FieldType.String; }
            else if (type == "Int") { fd_type = FieldType.Integer; }
            else if (type == "Date") { fd_type = FieldType.Date; }
            return fd_type;
        }

        // 字段类型转换
        public static string GetFeildType(string type, bool isToDouble)
        {
            string fd_type = type;
            if (type == "Float" && isToDouble) { fd_type = "DOUBLE"; }
            else if (type == "Char" || type == "VarChar" || type == "String") { fd_type = "TEXT"; }
            else if (type == "Int") { fd_type = "SHORT"; }
            else if (type == "Date") { fd_type = "DATE"; }
            return fd_type;
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://mp.weixin.qq.com/s?__biz=Mzg2MTc0MjIwNA==&mid=2247487637&idx=1&sn=cf238bc6f8f262b3b52f7597876d5e76&chksm=ce132586f964ac90a0c8ac677e5f95b6ee18ab49159e72c285e2f096e9e1efd49e42ea9eda0b#rd";
            UITool.Link2Web(url);
        }

        private void downLoadButton_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://pan.baidu.com/s/1-UkLi0gkNlIgfrYYIm8DSQ?pwd=6r3d";
            UITool.Link2Web(url);
        }

    }
}
