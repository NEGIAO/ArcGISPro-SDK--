using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.Util;
using System;
using System.Collections.Generic;
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
    /// Interaction logic for WriteGDToExcel.xaml
    /// </summary>
    public partial class WriteGDToExcel : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "WriteGDToExcel";

        public WriteGDToExcel()
        {
            InitializeComponent();

            // 初始化其它参数选项
            textFolderPath.Text = BaseTool.ReadValueFromReg(toolSet, "folderPath");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "国有耕地摸底排查表";


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
                string folderPath = textFolderPath.Text;

                // 判断参数是否选择完全
                if (sd == "" || folderPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "folderPath", folderPath);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(sd);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    //  读取图斑
                    FeatureLayer featureLayer = sd.TargetFeatureLayer();

                    // 图斑个数
                    long featureCount = featureLayer.GetFeatureCount();

                    using RowCursor rowCursor = featureLayer.Search();
                    while (rowCursor.MoveNext())
                    {
                        using Feature feature = rowCursor.Current as Feature;
                        // 获取字段值
                        string BSM = feature["BSM"]?.ToString();
                        string ZLDWMC = feature["ZLDWMC"]?.ToString();
                        string BZ = feature["BZ"]?.ToString();
                        string XJMC = feature["XJMC"]?.ToString();
                        string ZJMC = feature["ZJMC"]?.ToString();

                        string YJDKBH = feature["YJDKBH"]?.ToString();
                        string TQTBDLMJ = feature["TQTBDLMJ"]?.ToString();
                        string XZGDMJ = feature["XZGDMJ"]?.ToString();
                        string SYQFMC = feature["SYQFMC"]?.ToString();
                        string SYQFCJ = feature["SYQFCJ"]?.ToString();

                        string GDSZLX = feature["GDSZLX"]?.ToString();
                        string GDXCSJ = feature["GDXCSJ"]?.ToString();
                        string XZSJDLMC = feature["XZSJDLMC"]?.ToString();
                        string SJSYFMC = feature["SJSYFMC"]?.ToString();
                        string SJSYFZJH = feature["SJSYFZJH"]?.ToString();

                        string SJSYFLXDH = feature["SJSYFLXDH"]?.ToString();
                        string SYLY = feature["SYLY"]?.ToString();
                        string SJZZFMC = feature["SJZZFMC"]?.ToString();
                        string SJZZFZJH = feature["SJZZFZJH"]?.ToString();
                        string SJZZFLXDH = feature["SJZZFLXDH"]?.ToString();

                        // 复制excel表格
                        string excelPath = $@"{folderPath}\地块_{BSM}.xlsx";
                        DirTool.CopyResourceFile(@"CCTool.Data.Excel.规划.国有耕地摸底排查表(市局).xlsx", excelPath);

                        pw.AddMessageMiddle(80 / featureCount, $"写入地块_{BSM}", Brushes.Gray);
                        // 获取工作薄、工作表
                        string excelFile = ExcelTool.GetPath(excelPath);
                        int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
                        // 打开工作薄
                        Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                        // 打开工作表
                        Worksheet sheet = wb.Worksheets[sheetIndex];
                        // 获取cell
                        Cells cells = sheet.Cells;

                        // 赋值
                        cells[2, 0].Value = $"排查单位：{ZLDWMC}";
                        cells[2, 4].Value = BZ;
                        cells[3, 1].Value = XJMC;
                        cells[3, 3].Value = ZJMC;
                        cells[3, 7].Value = ZLDWMC;

                        cells[4, 1].Value = YJDKBH;
                        cells[4, 3].Value = TQTBDLMJ;
                        cells[4, 7].Value = XZGDMJ;

                        cells[5, 1].Value = SYQFMC;
                        cells[5, 6].Value = SYQFCJ;
                        cells[6, 1].Value = GDSZLX;

                        cells[7, 1].Value = GDXCSJ;
                        cells[7, 7].Value = XZSJDLMC;
                        cells[9, 2].Value = SJSYFMC;
                        cells[9, 6].Value = SJSYFZJH;
                        cells[9, 8].Value = SJSYFLXDH;

                        cells[10, 1].Value = SYLY;
                        cells[12, 2].Value = SJZZFMC;
                        cells[12, 6].Value = SJZZFZJH;
                        cells[12, 8].Value = SJZZFLXDH;

                        // 保存
                        wb.Save(excelFile);
                        wb.Dispose();
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


        private void combox_sd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_sd);
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textFolderPath.Text = UITool.OpenDialogFolder();
        }

        private List<string> CheckData(string sd)
        {
            List<string> result = new List<string>();

            // 检查是否有指定字段值
            List<string> fileds = new List<string>() { "ZLDWDM", "ZLDWMC" };
            string fieldResult = CheckTool.CheckFieldValueEmpty(sd, fileds);
            if (fieldResult != "")
            {
                result.Add(fieldResult);
            }
            return result;
        }

    }
}
