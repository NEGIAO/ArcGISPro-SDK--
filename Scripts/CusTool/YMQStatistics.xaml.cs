using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
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
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CCTool.Scripts.CusTool
{

    /// <summary>
    /// Interaction logic for YMQStatistics.xaml
    /// </summary>
    public partial class YMQStatistics : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public YMQStatistics()
        {
            InitializeComponent();
            combox_areaType.Items.Add("投影面积");
            combox_areaType.Items.Add("椭球面积");
            combox_areaType.SelectedIndex = 0;

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "淹没区分析(BHM)";


        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                var init_gdb = Project.Current.DefaultGeodatabasePath;
                // 获取分析要素
                string gd = combox_gd.ComboxText();
                string jbnt = combox_jbnt.ComboxText();
                string ft = combox_ft.ComboxText();
                string road = combox_road.ComboxText();

                // 获取字段
                string gd_field1 = combox_gd_field1.ComboxText();
                string gd_field2 = combox_gd_field2.ComboxText();
                string jbnt_field1 = combox_jbnt_field1.ComboxText();
                string jbnt_field2 = combox_jbnt_field2.ComboxText();
                string ft_field = combox_ft_field.ComboxText();
                string road_field = combox_road_field.ComboxText();

                // 获取分区
                string zone = combox_zone.ComboxText();

                string excel_path = textExcelPath.Text;

                string areaType = combox_areaType.Text;

                // 判断参数是否选择完全
                if (zone == "" || gd == "" || jbnt == "" || ft == "" || road == "" || gd_field1 == "" || gd_field2 == ""
                    || ft_field == "" || road_field == "" || excel_path == "")
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
                    // 复制嵌入资源中的Excel文件
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.淹没区统计表.xlsx", excel_path);

                    pw.AddMessageStart("汇总各要素指标");
                    Dictionary<string, string> featureLayers = new Dictionary<string, string>()
                    {
                        { gd, $"{gd_field1};{gd_field2}"},
                        { jbnt, $"{gd_field1};{gd_field2}"},
                        { ft, ft_field},
                        { road, road_field},
                    };
                    
                    string ly_clip = @$"{init_gdb}\ly_clip";
                    // 汇总各要素指标
                    foreach (var fl in featureLayers)
                    {
                        // 裁剪
                        Arcpy.Clip(fl.Key, zone, ly_clip);
                        // 添加一个面积字段
                        Arcpy.AddField(ly_clip, "mj_sx", "DOUBLE");
                        if (fl.Key != road)
                        {
                            // 计算投影或椭球面积
                            if (areaType == "投影面积")
                            {
                                Arcpy.CalculateField(ly_clip, "mj_sx", "!shape.area!");
                            }
                            else
                            {
                                Arcpy.CalculateField(ly_clip, "mj_sx", "!shape.geodesicarea!");
                            }
                        }
                        else
                        {
                            Arcpy.CalculateField(ly_clip, "mj_sx", "!shape.length!");
                        }
                        
                        // 汇总面积
                        Arcpy.Statistics(ly_clip, @$"{init_gdb}\{fl.Key}_table", "mj_sx SUM", fl.Value);
                    }

                    // 打开工作薄
                    Workbook wb = ExcelTool.OpenWorkbook(excel_path);
                    pw.AddMessageMiddle(30, "获取指标并写入");
                    pw.AddMessageMiddle(30, "耕地", Brushes.Gray);
                    // 打开工作表
                    Worksheet worksheet = wb.Worksheets["耕地"];
                    // 获取Cells
                    Cells cells = worksheet.Cells;
                    wb.CalculateFormula();
                    // 获取指标并写入
                    Dictionary<List<string>, double> dict_gd = GisTool.GetDictListFromPathDouble(@$"{init_gdb}\{gd}_table", new List<string>() { gd_field1, gd_field2 }, "SUM_mj_sx");
                    // 写入指标
                    int index = 2;
                    foreach (var gd in dict_gd)
                    {
                        if (index > 3) { cells.InsertRow(index); }// 如果不是第一行，复制行
                        cells[index, 0].PutValue( index -1);
                        cells[index, 1].PutValue(gd.Key[0]);
                        cells[index, 2].PutValue(gd.Key[1]);
                        cells[index, 3].PutValue(gd.Value);
                        index++;
                    }
                    // 写入统计字段
                    cells[1, 1].PutValue(gd_field1);
                    cells[1, 2].PutValue(gd_field2);
                    cells[index, 3].Formula =$"=SUM(D3:D{index})";    // 写入统计值

                    pw.AddMessageMiddle(10, "永农", Brushes.Gray);
                    // 打开工作表
                    worksheet = wb.Worksheets["永农"];
                    // 获取Cells
                    cells = worksheet.Cells;
                    // 获取指标并写入
                    Dictionary<List<string>, double> dict_jbnt = GisTool.GetDictListFromPathDouble(@$"{init_gdb}\{jbnt}_table", new List<string>() { jbnt_field1, jbnt_field2 }, "SUM_mj_sx");
                    // 写入指标
                    index = 2;
                    foreach (var jbnt in dict_jbnt)
                    {
                        if (index > 3) { cells.InsertRow(index); }// 如果不是第一行，复制行
                        cells[index, 0].PutValue(index - 1);
                        cells[index, 1].PutValue(jbnt.Key[0]);
                        cells[index, 2].PutValue(jbnt.Key[1]);
                        cells[index, 3].PutValue(jbnt.Value);
                        index++;
                    }
                    // 写入统计字段
                    cells[1, 1].PutValue(jbnt_field1);
                    cells[1, 2].PutValue(jbnt_field2);
                    cells[index, 3].Formula = $"=SUM(D3:D{index})";    // 写入统计值

                    pw.AddMessageMiddle(10, "房台", Brushes.Gray);
                    // 打开工作表
                    worksheet = wb.Worksheets["房台"];
                    // 获取Cells
                    cells = worksheet.Cells;
                    // 获取指标并写入
                    Dictionary<string, double> dict_ft= GisTool.GetDictFromPathDouble(@$"{init_gdb}\{ft}_table",  ft_field , "SUM_mj_sx");
                    // 写入指标
                    index = 2;
                    foreach (var ft in dict_ft)
                    {
                        if (index > 3) { cells.InsertRow(index); }// 如果不是第一行，复制行
                        cells[index, 0].PutValue(index - 1);
                        cells[index, 1].PutValue(ft.Key);
                        cells[index, 2].PutValue(ft.Value);
                        index++;
                    }
                    // 写入统计字段
                    cells[1, 1].PutValue(ft_field);
                    cells[index, 2].Formula = $"=SUM(C3:C{index})";    // 写入统计值

                    pw.AddMessageMiddle(10, "道路", Brushes.Gray);
                    // 打开工作表
                    worksheet = wb.Worksheets["道路"];
                    // 获取Cells
                    cells = worksheet.Cells;
                    // 获取指标并写入
                    Dictionary<string, double> dict_road = GisTool.GetDictFromPathDouble(@$"{init_gdb}\{road}_table", road_field, "SUM_mj_sx");
                    // 写入指标
                    index = 2;
                    foreach (var road in dict_road)
                    {
                        if (index > 3) { cells.InsertRow(index); }// 如果不是第一行，复制行
                        cells[index, 0].PutValue(index - 1);
                        cells[index, 1].PutValue(road.Key);
                        cells[index, 2].PutValue(road.Value);
                        index++;
                    }
                    // 写入统计字段
                    cells[1, 1].PutValue(road_field);
                    cells[index, 2].Formula = $"=SUM(C3:C{index})";    // 写入统计值

                    // 保存
                    wb.Save(excel_path);
                    wb.Dispose();

                    // 删除中间数据
                    Arcpy.Delect(@$"{init_gdb}\gd_table");
                    Arcpy.Delect(@$"{init_gdb}\jbnt_table");
                    Arcpy.Delect(@$"{init_gdb}\ft_table");
                    Arcpy.Delect(@$"{init_gdb}\road_table");
                    Arcpy.Delect(ly_clip);

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
            string url = "https://blog.csdn.net/xcc34452366/article/details/139253712";
            UITool.Link2Web(url);
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.SaveDialogExcel();
        }

        private void combox_zone_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_zone);
        }

        private void combox_gd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_gd);
        }

        private void combox_jbnt_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_jbnt);
        }

        private void combox_ft_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_ft);
        }

        private void combox_road_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_road, "Polyline");
        }

        private void combox_gd_field1_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_gd.ComboxText(), combox_gd_field1);
        }

        private void combox_gd_field2_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_gd.ComboxText(), combox_gd_field2);
        }

        private void combox_jbnt_field1_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_jbnt.ComboxText(), combox_jbnt_field1);
        }

        private void combox_jbnt_field2_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_jbnt.ComboxText(), combox_jbnt_field2);
        }

        private void combox_ft_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_ft.ComboxText(), combox_ft_field);
        }

        private void combox_road_field_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_road.ComboxText(), combox_road_field);
        }
    }


}
