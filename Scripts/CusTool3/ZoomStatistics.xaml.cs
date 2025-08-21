using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using Aspose.Words.Fields;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
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
using FieldType = ArcGIS.Core.Data.FieldType;
using Row = ArcGIS.Core.Data.Row;
using Table = ArcGIS.Core.Data.Table;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for ZoomStatistics.xaml
    /// </summary>
    public partial class ZoomStatistics : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public ZoomStatistics()
        {
            InitializeComponent();

            // 初始化combox
            combox_unit.Items.Add("平方米");
            combox_unit.Items.Add("公顷");
            combox_unit.Items.Add("平方公里");
            combox_unit.Items.Add("亩");
            combox_unit.Items.Add("万亩");
            combox_unit.SelectedIndex = 1;

            combox_area.Items.Add("平面面积");
            combox_area.Items.Add("椭球面积");
            combox_area.SelectedIndex = 0;

            combox_digit.Items.Add("1");
            combox_digit.Items.Add("2");
            combox_digit.Items.Add("3");
            combox_digit.Items.Add("4");
            combox_digit.Items.Add("5");
            combox_digit.Items.Add("6");
            combox_digit.SelectedIndex = 1;

            // 工程默认文件夹位置
            string folder_path = Project.Current.HomeFolderPath;
            textExcelPath.Text = $@"{folder_path}\多图层统计结果.xlsx";
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "多图层分区统计(lan)";


        private void combox_zoom_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_zoom);
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.SaveDialogExcel();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string zoom = combox_zoom.ComboxText();
                string area_type = combox_area.Text[..2];
                string unit = combox_unit.Text;
                int digit = int.Parse(combox_digit.Text);
                string excel_path = textExcelPath.Text;

                // 获取参数listbox
                List<string> fieldNames = UITool.GetTextFromListBox(listbox_field);
                List<string> targetFeatureClasses = UITool.GetTextFromListBox(listbox_targetFeature);

                // 统计字段拼合
                string fieldString = "";
                foreach (string fieldName in fieldNames)
                {
                    fieldString += $"{fieldName};";
                }
                fieldString = fieldString[..^1];

                // 默认数据库位置
                var gdb_path = Project.Current.DefaultGeodatabasePath;
                // 工程默认文件夹位置
                string folder_path = Project.Current.HomeFolderPath;

                // 判断参数是否选择完全
                if (zoom == "" || excel_path == "" || fieldNames.Count == 0 || targetFeatureClasses.Count == 0)
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
                    // 检查数据
                    List<string> errs = CheckData(zoom, fieldNames);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }


                    pw.AddMessageMiddle(10, "初始化excel表格");

                    // 复制嵌入资源中的Excel文件
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.多图层统计模板.xlsx", excel_path);

                    // 先汇总统计一下所有的行
                    string zoom_table = $@"{gdb_path}\zoom_table";
                    string IDField = zoom.TargetIDFieldName();
                    Arcpy.Statistics(zoom, zoom_table, $"{IDField} SUM", fieldString);
                    // 写入行和列
                    ExcelWriteRows(excel_path, zoom_table, fieldNames, targetFeatureClasses, digit);

                    pw.AddMessageMiddle(20, "统计数据");

                    string Identity = $@"{gdb_path}\Identity";
                    string clip = $@"{gdb_path}\clip";
                    string Identity_table = $@"{gdb_path}\Identity_table";

                    foreach (string targetFeatureClass in targetFeatureClasses)
                    {
                        pw.AddMessageMiddle(5, $"       统计项：{targetFeatureClass}", Brushes.Gray);

                        // 剪裁
                        Arcpy.Clip(targetFeatureClass, zoom, clip);

                        // 标识
                        Arcpy.Identity(clip, zoom, Identity);
                        // 计算面积
                        GisTool.AddField(Identity, "计算面积", FieldType.Double);
                        if (area_type == "平面")
                        {
                            Arcpy.CalculateField(Identity, "计算面积", "!shape.area!");
                        }
                        else
                        {
                            Arcpy.CalculateField(Identity, "计算面积", "!shape.geodesicarea!");
                        }
                        // 汇总
                        Arcpy.Statistics(Identity, Identity_table, $"计算面积 SUM", fieldString);

                        // 写入统计项
                        ExcelWriteSta(excel_path, Identity_table, targetFeatureClass, fieldNames, unit);
                    }

                    // 删除中间数据
                    Arcpy.Delect(zoom_table);
                    Arcpy.Delect(Identity);
                    Arcpy.Delect(Identity_table);
                });

                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void ExcelWriteRows(string excel_path, string zoom_table, List<string> fieldNames, List<string> targetFeatureClasses, int digit)
        {
            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excel_path);
            int sheetIndex = ExcelTool.GetSheetIndex(excel_path);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取cells
            Cells cells = sheet.Cells;
            // 插入列
            for (int i = 0; i < targetFeatureClasses.Count; i++)
            {
                if (i > 0)
                {
                    cells.CopyColumn(cells, 1, i + 1);
                }
                cells[0, i + 1].Value = targetFeatureClasses[i];
            }

            for (int i = 0; i < fieldNames.Count; i++)
            {
                if (i > 0)
                {
                    cells.InsertColumn(i);
                }
                cells[0, i].Value = fieldNames[i];
            }

            // 写入行
            using Table table = zoom_table.TargetTable();
            using RowCursor rowCursor = table.Search();
            int index = 1;
            while (rowCursor.MoveNext())
            {
                // 复制行
                if (index > 1)
                {
                    cells.CopyRow(cells, 1, index);
                }

                using Row row = rowCursor.Current;
                List<string> valueList = new List<string>();
                foreach (string fieldName in fieldNames)
                {
                    valueList.Add(row[fieldName].ToString());
                }
                if (valueList[0] != "")
                {
                    // 写入
                    for (int i = 0; i < fieldNames.Count; i++)
                    {
                        cells[index, i].Value = valueList[i];
                    }
                }
                index++;
            }

            // 插入合计行
            int totalRow = cells.MaxDataRow+1;
            cells.CopyRow(cells, 1, totalRow);
            // 合计行修改
            cells.Merge(totalRow, 0, 1, fieldNames.Count);
            cells[totalRow, 0].Value = "合计";

            // 设置单元格为数字型，小数位数
            Aspose.Cells.Style style = cells[1, fieldNames.Count].GetStyle();
            style.Number = 4;   // 数字型
            style.Custom = digit switch
            {
                1 => "0.0",
                2 => "0.00",
                3 => "0.000",
                4 => "0.0000",
                _ => null,
            };
            // 设置
            for (int i = 1; i <= cells.MaxDataRow ; i++)
            {
                for (int j = fieldNames.Count; j <= cells.MaxColumn; j++)
                {
                    // 设置
                    cells[i, j].SetStyle(style);
                }
            }

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }


        private void ExcelWriteSta(string excel_path, string Identity_table, string targetFeatureClass, List<string> fieldNames, string unit)
        {
            // 单位系数设置
            double unit_xs = unit switch
            {
                "平方米" => 1,
                "公顷" => 10000,
                "平方公里" => 1000000,
                "亩" => 666.66667,
                "万亩" => 6666666.66667,
                _ => 1,
            };

            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excel_path);
            int sheetIndex = ExcelTool.GetSheetIndex(excel_path);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取cells
            Cells cells = sheet.Cells;

            // 找出指定列
            int initCol = 0;
            for (int i = 0; i <= cells.MaxDataColumn; i++)
            {
                if (cells[0, i].StringValue == targetFeatureClass)
                {
                    initCol = i;
                }
            }


            // 写入行
            using Table table = Identity_table.TargetTable();
            using RowCursor rowCursor = table.Search();
            int index = 1;
            while (rowCursor.MoveNext())
            {
                using Row row = rowCursor.Current;

                // 面积值
                _ = double.TryParse(row["SUM_计算面积"].ToString(), out double mj);

                // 当前行的值
                List<string> valueList = new();
                foreach (string fieldName in fieldNames)
                {
                    valueList.Add(row[fieldName].ToString());
                }
                // excel行的值
                for (int i = 0; i <= cells.MaxDataRow; i++)
                {
                    // 符号要求的行
                    int initRow = 0;

                    List<string> cellValues = new();
                    for (int j = 0; j < fieldNames.Count; j++)
                    {
                        cellValues.Add(cells[i, j].StringValue);
                    }
                    // 判断是否完全符合
                    bool isEqul = true;
                    for (int k = 0; k < cellValues.Count; k++)
                    {
                        if (cellValues[k] != valueList[k])
                        {
                            isEqul = false;
                            break;
                        }
                        initRow = i;
                    }

                    // 如果符号，就写入
                    if (isEqul)
                    {
                        cells[initRow, initCol].Value = mj / unit_xs;
                    }
                }

                index++;
            }

            // 合计值
            double total_mj = 0;
            for (int i = 1;i < cells.MaxDataRow;i++)
            {
                bool result = double.TryParse(cells[i, initCol].StringValue, out double mj);
                if (result)
                {
                    total_mj += mj;
                }
            }
            cells[cells.MaxDataRow, initCol].Value = total_mj;

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }


        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/144606103?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        private void combox_zoomField_DropClose(object sender, EventArgs e)
        {
            string zoomField = combox_zoomField.ComboxText();

            var list = UITool.GetTextFromListBox(listbox_field);

            if (!list.Contains(zoomField))
            {
                listbox_field.Items.Add(zoomField);
            }
        }

        private void combox_zoomField_DropDown(object sender, EventArgs e)
        {
            string zoom = combox_zoom.ComboxText();
            if (zoom != "")
            {
                UITool.AddTextFieldsToComboxPlus(combox_zoom.ComboxText(), combox_zoomField);
            }
        }

        private void combox_sta_DropClose(object sender, EventArgs e)
        {
            try
            {
                string sta_fc = combox_sta.ComboxText();

                var list = UITool.GetTextFromListBox(listbox_targetFeature);

                if (!list.Contains(sta_fc))
                {
                    listbox_targetFeature.Items.Add(sta_fc);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        private void combox_sta_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_sta);
        }

        private void btn_fcCanser_Click(object sender, RoutedEventArgs e)
        {
            // 获取ListBox的Items集合
            var items = listbox_targetFeature.Items;
            // 从Items集合中移除选定的项
            var selectedItem = listbox_targetFeature.SelectedItem;
            items.Remove(selectedItem);
        }

        private void btn_fieldCanser_Click(object sender, RoutedEventArgs e)
        {

            // 获取ListBox的Items集合
            var items = listbox_field.Items;
            // 从Items集合中移除选定的项
            var selectedItem = listbox_field.SelectedItem;
            items.Remove(selectedItem);

        }

        private void btn_fieldUp_Click(object sender, RoutedEventArgs e)
        {
            // 获取选定项的索引
            int selectedIndex = listbox_field.SelectedIndex;

            if (selectedIndex > 0)
            {
                // 获取ListBox的Items集合
                var items = listbox_field.Items;
                // 从Items集合中移除选定的项
                var selectedItem = listbox_field.SelectedItem;
                items.Remove(selectedItem);
                // 在前一个位置重新插入选定的项
                items.Insert(selectedIndex - 1, selectedItem);
                // 更新ListBox的选定项
                listbox_field.SelectedItem = selectedItem;
            }
        }

        private void btn_fieldDown_Click(object sender, RoutedEventArgs e)
        {
            // 获取选定项的索引
            int selectedIndex = listbox_field.SelectedIndex;

            if (selectedIndex < listbox_field.Items.Count - 1)
            {
                // 获取ListBox的Items集合
                var items = listbox_field.Items;
                // 从Items集合中移除选定的项
                var selectedItem = listbox_field.SelectedItem;
                items.Remove(selectedItem);
                // 在前一个位置重新插入选定的项
                items.Insert(selectedIndex + 1, selectedItem);
                // 更新ListBox的选定项
                listbox_field.SelectedItem = selectedItem;
            }
        }

        private void btn_fcUp_Click(object sender, RoutedEventArgs e)
        {
            // 获取选定项的索引
            int selectedIndex = listbox_targetFeature.SelectedIndex;

            if (selectedIndex > 0)
            {
                // 获取ListBox的Items集合
                var items = listbox_targetFeature.Items;
                // 从Items集合中移除选定的项
                var selectedItem = listbox_targetFeature.SelectedItem;
                items.Remove(selectedItem);
                // 在前一个位置重新插入选定的项
                items.Insert(selectedIndex - 1, selectedItem);
                // 更新ListBox的选定项
                listbox_targetFeature.SelectedItem = selectedItem;
            }
        }

        private void btn_fcDown_Click(object sender, RoutedEventArgs e)
        {
            // 获取选定项的索引
            int selectedIndex = listbox_targetFeature.SelectedIndex;

            if (selectedIndex < listbox_targetFeature.Items.Count - 1)
            {
                // 获取ListBox的Items集合
                var items = listbox_targetFeature.Items;
                // 从Items集合中移除选定的项
                var selectedItem = listbox_targetFeature.SelectedItem;
                items.Remove(selectedItem);
                // 在前一个位置重新插入选定的项
                items.Insert(selectedIndex + 1, selectedItem);
                // 更新ListBox的选定项
                listbox_targetFeature.SelectedItem = selectedItem;
            }
        }

        private List<string> CheckData(string fc, List<string> fieldNames)
        {
            List<string> result = new List<string>();

            // 检查字段值是否为空【ZLDWDM,ZLDWMC】
            string fieldEmptyResult = CheckTool.CheckFieldValueSpace(fc, fieldNames);
            if (fieldEmptyResult != "")
            {
                result.Add(fieldEmptyResult);
            }

            return result;
        }

    }
}
