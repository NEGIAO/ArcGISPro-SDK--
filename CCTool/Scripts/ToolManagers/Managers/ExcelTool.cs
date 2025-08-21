using Aspose.Cells;
using Aspose.Cells.Rendering;
using CCTool.Scripts.ToolManagers.Extensions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Managers
{
    public class ExcelTool
    {
        // 打开工作薄
        public static Workbook OpenWorkbook(string excelFile)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            LoadOptions loadOptinos = new LoadOptions(LoadFormat.Auto);
            Workbook wb = new Workbook(excelFile, loadOptinos);
            return wb;
        }

        // 获取Excel的文件名
        public static string GetPath(string excelPath)
        {
            // 如果是完整路径（包含sheet$）
            if (excelPath.Contains('$'))
            {
                // 获取最后一个"\"的位置
                int index = excelPath.LastIndexOf("\\");
                // 获取exl文件名
                string excel_name = excelPath[..index];
                // 返回exl文件名
                return excel_name;
            }
            // 如果只excel文件路径
            else
            {
                return excelPath;
            }
        }

        // 获取Excel的表序号
        public static int GetSheetIndex(string excelPath)
        {
            // 如果是完整路径（包含sheet$）
            if (excelPath.Contains('$'))
            {
                // 获取最后一个"\"的位置
                int index = excelPath.LastIndexOf("\\");
                // 获取exl文件名
                string excelName = excelPath[..index];
                // 获取表名
                string sheet_name = excelPath.Substring(index + 1, excelPath.Length - index - 2);
                // 打开工作薄
                Workbook wb = OpenWorkbook(excelName);
                // 获取第一个工作表
                Worksheet sheet = wb.Worksheets[sheet_name];
                //  返回index
                return sheet.Index;
            }
            // 如果只excel文件路径
            else
            {
                return 0;
            }
        }


        // Excel文件属性映射【输入映射字典dict】
        public static void AttributeMapper(string excelPath, int sheet_in_col, int sheet_map_col, Dictionary<string, string> dict, int startRow = 0)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 逐行处理
            for (int i = startRow; i <= sheet.Cells.MaxDataRow; i++)
            {
                //  获取目标cell
                Cell inCell = sheet.Cells[i, sheet_in_col];
                Cell mapCell = sheet.Cells[i, sheet_map_col];
                // 属性映射
                if (inCell is not null && dict.ContainsKey(inCell.StringValue))
                {
                    mapCell.Value = dict[inCell.StringValue];   // 赋值
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }


        // Excel文件属性映射【输入映射字典dict】
        public static void AttributeMapperDouble(string excelPath, int sheet_in_col, int sheet_map_col, Dictionary<string, double> dict, int startRow = 0)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 逐行处理
            for (int i = startRow; i <= sheet.Cells.MaxDataRow; i++)
            {
                //  获取目标cell
                Cell inCell = sheet.Cells[i, sheet_in_col];
                Cell mapCell = sheet.Cells[i, sheet_map_col];
                // 属性映射
                if (inCell is not null && dict.ContainsKey(inCell.StringValue))
                {
                    mapCell.Value = dict[inCell.StringValue];   // 赋值
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel文件属性映射【输入映射字典dict】
        public static void AttributeMapperDouble(string excelPath, int sheet_in_col, int sheet_map_col, Dictionary<string, string> dict, int startRow = 0)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 逐行处理
            for (int i = startRow; i <= sheet.Cells.MaxDataRow; i++)
            {
                //  获取目标cell
                Cell inCell = sheet.Cells[i, sheet_in_col];
                Cell mapCell = sheet.Cells[i, sheet_map_col];
                // 属性映射
                if (inCell is not null && dict.ContainsKey(inCell.StringValue))
                {
                    mapCell.Value = double.Parse(dict[inCell.StringValue]);   // 赋值
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel文件属性映射_列【输入映射字典dict】
        public static void AttributeMapperCol(string excelPath, int sheet_in_row, int sheet_map_row, Dictionary<string, string> dict, int startCol = 0)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 逐行处理
            for (int i = startCol; i <= sheet.Cells.MaxDataColumn; i++)
            {
                //  获取目标cell
                Cell inCell = sheet.Cells[sheet_in_row, i];
                Cell mapCell = sheet.Cells[sheet_map_row, i];
                // 属性映射
                if (inCell is not null && dict.ContainsKey(inCell.StringValue))
                {
                    mapCell.Value = dict[inCell.StringValue];   // 赋值
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel文件属性映射_列【输入映射字典dict】
        public static void AttributeMapperColDouble(string excelPath, int sheet_in_row, int sheet_map_row, Dictionary<string, double> dict, int startCol = 0)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 逐行处理
            for (int i = startCol; i <= sheet.Cells.MaxDataColumn; i++)
            {
                //  获取目标cell
                Cell inCell = sheet.Cells[sheet_in_row, i];
                Cell mapCell = sheet.Cells[sheet_map_row, i];
                // 属性映射
                if (inCell is not null && dict.ContainsKey(inCell.StringValue))
                {
                    mapCell.Value = dict[inCell.StringValue];   // 赋值
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // 复制Excel表中sheet
        public static void CopySheet(string excelPath, string oldSheet, string newSheet)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);

            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 旧表
            Worksheet oldWS = wb.Worksheets[oldSheet];

            // 创建新表
            int newSheetIndex = wb.Worksheets.Add();
            // 新表
            Worksheet newWS = wb.Worksheets[newSheetIndex];
            // 设置新表的名称
            newWS.Name = newSheet;
            // 复制
            newWS.Copy(oldWS);

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // 删除Excel表中sheet
        public static void DeleteSheet(string excelPath, string sheetName)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);

            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);

            // 移除sheet
            wb.Worksheets.RemoveAt(sheetName);

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel文件cell写入
        public static void WriteCell(string excelPath, int row, int col, string cell_value)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取cell
            Cell cell = sheet.Cells[row, col];
            // 写入cell值
            cell.Value = cell_value;
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel文件cell写入
        public static void WriteCell(string excelPath, int row, int col, double cell_value)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取cell
            Cell cell = sheet.Cells[row, col];
            // 写入cell值
            cell.Value = cell_value;
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel设置列单元格的格式
        public static void SetColStyle(string excelPath, int col, int startRow, int styleNumber = 4, int digit = 2)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];

            for (int i = startRow; i <= sheet.Cells.MaxDataRow; i++)
            {
                // 获取cell
                Cell cell = sheet.Cells[i, col];
                // 获取style
                Style style = cell.GetStyle();
                // 数字型
                style.Number = styleNumber;
                // 小数位数
                if (digit == 1) { style.Custom = "0.0"; }
                else if (digit == 2) { style.Custom = "0.00"; }
                else if (digit == 3) { style.Custom = "0.000"; }
                else if (digit == 4) { style.Custom = "0.0000"; }
            }

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }


        // Excel文件删除行(无合并格的情况)
        public static void DelectRowSimple(string excelPath, int row)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 删除行
            sheet.Cells.DeleteRow(row);

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel文件删除行(无合并格的情况)
        public static void DelectRowSimple(string excelPath, List<int> rows)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 删除行
            foreach (var row in rows)
            {
                sheet.Cells.DeleteRow(row);
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel文件删除列(无合并格的情况)
        public static void DelectColSimple(string excelPath, int col)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 删除列
            sheet.Cells.DeleteColumn(col);
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel文件删除多列(无合并格的情况)
        public static void DelectColSimple(string excelPath, List<int> cols)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 删除列
            foreach (var col in cols)
            {
                sheet.Cells.DeleteColumn(col);
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // 删除sheet表中的所有合并格，并填充默认值
        private static List<CellRangeAddress> RemoveMerge(string excelPath)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 创建文件流
            FileStream fs = File.Open(excelFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            // 打开工作薄
            XSSFWorkbook wb = new XSSFWorkbook(fs);
            // 获取工作表
            ISheet sheet = wb.GetSheetAt(sheetIndex);
            // 获取所有合并区域
            List<CellRangeAddress> mergeRanges = sheet.MergedRegions;
            // 设置一个List<CellRangeAddress>
            List<CellRangeAddress> upDataRanges = new List<CellRangeAddress>();
            // 检查并清除合并区域
            if (mergeRanges.Count > 0)
            {
                for (int i = mergeRanges.Count - 1; i >= 0; i--)
                {
                    // 合并格的四至
                    CellRangeAddress region = mergeRanges[i];
                    int firstRow = region.FirstRow;
                    int lastRow = region.LastRow;
                    int firstCol = region.FirstColumn;
                    int lastCol = region.LastColumn;
                    // 判定要处理的区域
                    for (int row = firstRow; row <= lastRow; row++)
                    {
                        for (int col = firstCol; col <= lastCol; col++)
                        {
                            if (row != firstRow || col != firstCol)
                            {
                                IRow r = sheet.GetRow(row);
                                ICell c = r.GetCell(col);
                                // 如果c是空值，则赋一个默认值
                                c ??= r.CreateCell(col);
                                // 设置拥有合并区域的单元格的值为合并区域的值
                                ICell mergedCell = sheet.GetRow(firstRow).GetCell(firstCol);

                                if (mergedCell != null)
                                {
                                    c.SetCellValue(mergedCell.StringCellValue); // 可根据需要选择相应的数据类型
                                }
                            }
                        }
                    }
                    // 计入
                    upDataRanges.Add(region);
                    // 清除合并区域
                    sheet.RemoveMergedRegion(i);
                }
            }
            // 保存工作簿
            wb.Write(new FileStream(excelFile, FileMode.Create, FileAccess.Write));
            // 返回
            return upDataRanges;
        }

        // 合并单元格（依据输入的CellRangeAddress和delectRowrow进行判断）
        private static void MergeFromAddressRow(string excelPath, List<CellRangeAddress> mergeRanges, int delectRow)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 创建文件流
            FileStream fs = File.Open(excelFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            // 打开工作薄
            XSSFWorkbook wb = new XSSFWorkbook(fs);
            // 获取工作表
            ISheet sheet = wb.GetSheetAt(sheetIndex);
            // 检查并更新合并区域
            foreach (CellRangeAddress mergeRange in mergeRanges)
            {
                // 如果合并格只有一行，且就是删除行，则不纳入合并的处理范围
                if (!(mergeRange.LastRow == mergeRange.FirstRow && mergeRange.LastRow == delectRow))
                {
                    // 如果合并格删除行的影响范围内
                    if (delectRow <= mergeRange.LastRow)
                    {
                        mergeRange.LastRow -= 1;
                    }
                    if (delectRow < mergeRange.FirstRow)
                    {
                        mergeRange.FirstRow -= 1;
                    }
                    // 重新合并单元格   判断合并单元格的格子数，不是单格才合并
                    if (mergeRange.NumberOfCells > 1)
                    {
                        sheet.AddMergedRegion(mergeRange);
                    }
                }
            }
            // 保存工作簿
            wb.Write(new FileStream(excelFile, FileMode.Create, FileAccess.Write));
        }

        // 合并单元格（依据输入的CellRangeAddress和delectRowrow进行判断）
        private static void MergeFromAddressRow(string excelPath, List<CellRangeAddress> mergeRanges, List<int> delectRows)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 创建文件流
            FileStream fs = File.Open(excelFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            // 打开工作薄
            XSSFWorkbook wb = new XSSFWorkbook(fs);
            // 获取工作表
            ISheet sheet = wb.GetSheetAt(sheetIndex);
            // 检查并更新合并区域
            foreach (CellRangeAddress mergeRange in mergeRanges)
            {
                bool isOver = false;
                foreach (int delectRow in delectRows)
                {
                    // 如果合并格只有一行，且就是删除行，则不纳入合并的处理范围
                    if (!(mergeRange.LastRow == mergeRange.FirstRow && mergeRange.LastRow == delectRow))
                    {
                        // 如果合并格删除行的影响范围内
                        if (delectRow <= mergeRange.LastRow)
                        {
                            mergeRange.LastRow -= 1;
                        }
                        if (delectRow < mergeRange.FirstRow)
                        {
                            mergeRange.FirstRow -= 1;
                        }
                    }
                    else
                    {
                        isOver = true;
                    }
                }
                // 重新合并单元格   判断合并单元格的格子数，不是单格才合并
                if (isOver == false && mergeRange.NumberOfCells > 1)
                {
                    sheet.AddMergedRegion(mergeRange);
                }
            }

            // 保存工作簿
            wb.Write(new FileStream(excelFile, FileMode.Create, FileAccess.Write));
        }

        // 合并单元格（依据输入的CellRangeAddress和delectRowrow进行判断）
        private static void MergeFromAddressCol(string excelPath, List<CellRangeAddress> mergeRanges, int delectCol)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 创建文件流
            FileStream fs = File.Open(excelFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            // 打开工作薄
            XSSFWorkbook wb = new XSSFWorkbook(fs);
            // 获取工作表
            ISheet sheet = wb.GetSheetAt(sheetIndex);
            // 检查并更新合并区域
            foreach (CellRangeAddress mergeRange in mergeRanges)
            {
                // 如果合并格只有一行，且就是删除行，则不纳入合并的处理范围
                if (!(mergeRange.LastColumn == mergeRange.FirstColumn && mergeRange.LastColumn == delectCol))
                {
                    // 如果合并格删除行的影响范围内
                    if (delectCol <= mergeRange.LastColumn)
                    {
                        mergeRange.LastColumn -= 1;
                    }
                    if (delectCol < mergeRange.FirstColumn)
                    {
                        mergeRange.FirstColumn -= 1;
                    }
                    // 重新合并单元格   判断合并单元格的格子数，不是单格才合并
                    if (mergeRange.NumberOfCells > 1)
                    {
                        sheet.AddMergedRegion(mergeRange);
                    }
                }
            }
            // 保存工作簿
            wb.Write(new FileStream(excelFile, FileMode.Create, FileAccess.Write));
        }

        // 合并单元格（依据输入的CellRangeAddress和delectRowrow进行判断）
        private static void MergeFromAddressCol(string excelPath, List<CellRangeAddress> mergeRanges, List<int> delectCols)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 创建文件流
            FileStream fs = File.Open(excelFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            // 打开工作薄
            XSSFWorkbook wb = new XSSFWorkbook(fs);
            // 获取工作表
            ISheet sheet = wb.GetSheetAt(sheetIndex);
            // 检查并更新合并区域
            foreach (CellRangeAddress mergeRange in mergeRanges)
            {
                bool isOver = false;
                // 循环每个删除列
                foreach (var delectCol in delectCols)
                {
                    // 如果合并格只有一行，且就是删除行，则不纳入合并的处理范围
                    if (!(mergeRange.LastColumn == mergeRange.FirstColumn && mergeRange.LastColumn == delectCol))
                    {
                        // 如果合并格删除行的影响范围内
                        if (delectCol <= mergeRange.LastColumn)
                        {
                            mergeRange.LastColumn -= 1;
                        }
                        if (delectCol < mergeRange.FirstColumn)
                        {
                            mergeRange.FirstColumn -= 1;
                        }
                    }
                    else
                    {
                        isOver = true;
                    }
                }
                // 重新合并单元格   判断合并单元格的格子数，不是单格才合并
                if (isOver == false && mergeRange.NumberOfCells > 1)
                {
                    sheet.AddMergedRegion(mergeRange);
                }
            }
            // 保存工作簿
            wb.Write(new FileStream(excelFile, FileMode.Create, FileAccess.Write));
        }

        // Excel文件删除行
        public static void DeleteRow(string excelPath, int delectRow)
        {
            // 删除sheet表中的所有合并格，并填充默认值
            List<CellRangeAddress> mergeRanges = RemoveMerge(excelPath);
            // 删除行
            DelectRowSimple(excelPath, delectRow);
            // 合并单元格（依据输入的CellRangeAddress和delectRowrow进行判断）
            MergeFromAddressRow(excelPath, mergeRanges, delectRow);
        }

        // Excel文件删除多行
        public static void DeleteRow(string excelPath, List<int> delectRows)
        {
            // 删除sheet表中的所有合并格，并填充默认值
            List<CellRangeAddress> mergeRanges = RemoveMerge(excelPath);
            // 删除多行
            DelectRowSimple(excelPath, delectRows);
            // 合并单元格（依据输入的CellRangeAddress和delectCol进行判断）
            MergeFromAddressRow(excelPath, mergeRanges, delectRows);
        }

        // Excel文件删除列
        public static void DeleteCol(string excelPath, int delectCol)
        {
            // 删除sheet表中的所有合并格，并填充默认值
            List<CellRangeAddress> mergeRanges = RemoveMerge(excelPath);
            // 删除列
            DelectColSimple(excelPath, delectCol);
            // 合并单元格（依据输入的CellRangeAddress和delectCol进行判断）
            MergeFromAddressCol(excelPath, mergeRanges, delectCol);
        }

        // Excel文件删除多列
        public static void DeleteCol(string excelPath, List<int> delectCols)
        {
            // 删除sheet表中的所有合并格，并填充默认值
            List<CellRangeAddress> mergeRanges = RemoveMerge(excelPath);
            // 删除多列
            DelectColSimple(excelPath, delectCols);
            // 合并单元格（依据输入的CellRangeAddress和delectCol进行判断）
            MergeFromAddressCol(excelPath, mergeRanges, delectCols);
        }

        // 删除Excel表中的0值行【指定1个列】
        private static List<int> DeleteNullRowResult(string excelPath, int deleteCol, int startRow = 0)
        {
            List<int> list = new List<int>();
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];

            // 强制更新表内的公式单元格
            wb.CalculateFormula();

            // 找出0值行
            for (int i = sheet.Cells.MaxDataRow; i >= startRow; i--)
            {
                Cell cell = sheet.Cells.GetCell(i, deleteCol);
                if (cell == null)  // 值为空则纳入
                {
                    list.Add(i);
                }
                else
                {
                    string str = cell.StringValue;
                    if (str == "")  // 值为0也纳入
                    {
                        list.Add(i);
                    }
                    else if (double.Parse(str) == 0)
                    {
                        list.Add(i);
                    }
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();

            // 返回值
            return list;
        }

        // 删除Excel表中的0值行【指定多个列】
        private static List<int> DeleteNullRowResult(string excelPath, List<int> deleteCols, int startRow = 0)
        {
            List<int> list = new List<int>();
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];

            // 强制更新表内的公式单元格
            wb.CalculateFormula();

            // 找出0值行
            for (int i = sheet.Cells.MaxDataRow; i >= startRow; i--)
            {
                // 设置一个flag
                bool isNull = true;
                // 循环查找各列的值
                foreach (var deleteCol in deleteCols)
                {
                    Cell cell = sheet.Cells.GetCell(i, deleteCol);
                    if (cell != null)  // 值不为空
                    {
                        string str = cell.StringValue;
                        if (str != "") // 值不为0
                        {
                            if (double.Parse(str) != 0)
                            {
                                isNull = false;
                                break;
                            }
                        }
                    }
                }
                // 输出删除列
                if (isNull)
                {
                    list.Add(i);
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
            // 返回值
            return list;
        }

        // 删除Excel表中的0值行【指定1个列】
        private static List<int> DeleteNullColResult(string excelPath, int deleteRow, int startCol = 0, int lastCol = 0)
        {
            List<int> list = new List<int>();
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];

            // 强制更新表内的公式单元格
            wb.CalculateFormula();

            // 找出0值行
            for (int i = sheet.Cells.MaxDataColumn + lastCol; i >= startCol; i--)
            {
                Cell cell = sheet.Cells.GetCell(deleteRow, i);
                if (cell == null)  // 值为空则纳入
                {
                    list.Add(i);
                }
                else
                {
                    string str = cell.StringValue;
                    if (str == "")  // 值为0也纳入
                    {
                        list.Add(i);
                    }
                    else if (double.Parse(str) == 0)
                    {
                        list.Add(i);
                    }
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();

            // 返回值
            return list;
        }

        // 删除Excel表中的0值行【指定多个列】
        private static List<int> DeleteNullColResult(string excelPath, List<int> deleteRows, int startCol = 0, int lastCol = 0)
        {
            List<int> list = new List<int>();
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];

            // 强制更新表内的公式单元格
            wb.CalculateFormula();

            int deleteCount = sheet.Cells.MaxDataColumn;

            // 找出0值行
            for (int i = deleteCount - lastCol; i >= startCol; i--)
            {
                // 设置一个flag
                bool isNull = true;
                // 循环查找各列的值
                foreach (var deleteRow in deleteRows)
                {
                    Cell cell = sheet.Cells.GetCell(deleteRow, i);
                    if (cell != null)  // 值不为空
                    {
                        string str = cell.StringValue;
                        if (str != "") // 值不为0
                        {
                            if (double.Parse(str) != 0)
                            {
                                isNull = false;
                                break;
                            }
                        }
                    }
                }
                // 输出删除列
                if (isNull)
                {
                    list.Add(i);
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();

            // 返回值
            return list;
        }

        // 删除Excel表中的0值行【指定1个列】
        public static void DeleteNullRow(string excelPath, int deleteCol, int startRow = 0)
        {
            // 要删除行
            List<int> deleleRows = DeleteNullRowResult(excelPath, deleteCol, startRow);
            // 删除行
            DeleteRow(excelPath, deleleRows);
        }

        // 删除Excel表中的0值行【指定多个列】
        public static void DeleteNullRow(string excelPath, List<int> deleteCols, int startRow = 0)
        {
            // 要删除行
            List<int> deleleRows = DeleteNullRowResult(excelPath, deleteCols, startRow);
            // 删除行
            DeleteRow(excelPath, deleleRows);
        }

        // 删除Excel表中的0值列【指定1个列】
        public static void DeleteNullCol(string excelPath, int deleteRow, int startCol = 0, int lastCol = 0)
        {
            // 要删除列
            List<int> deleleRows = DeleteNullColResult(excelPath, deleteRow, startCol, lastCol);
            // 删除列
            DeleteCol(excelPath, deleleRows);
        }

        // 删除Excel表中的0值列【指定多个列】
        public static void DeleteNullCol(string excelPath, List<int> deleteRows, int startCol = 0, int lastCol = 0)
        {
            // 要删除列
            List<int> deleleCols = DeleteNullColResult(excelPath, deleteRows, startCol, lastCol);

            // 删除列
            DeleteCol(excelPath, deleleCols);
        }


        // 单元格合计【列】
        public static void StatisticsColCell(string excelPath, int col, int  startRow, int lastRow, int totalRow)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];

            Cells cells = sheet.Cells;

            double total = 0;

            for (int i = startRow; i <= lastRow; i++)
            {
                string str = cells[i, col].StringValue;
                if (str != "" && str is not null)
                {
                    total += cells[i, col].DoubleValue;
                }
                
            }

            // 赋值
            cells[totalRow, col].Value = total;
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // 复制Excel表中的行
        public static void CopyRows(string excelPath, int sourceRowIndex, int targetRowIndex, int count = 1)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];

            Cells cells = sheet.Cells;
            // 复制
            cells.CopyRows(cells, sourceRowIndex, targetRowIndex, count);
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // 从Excel文件中获取表头列表
        public static List<string> GetColListFromExcel(string excelPath)
        {
            // 定义列表
            List<string> result = new List<string>();
            List<string> stringList1 = new List<string>();
            List<string> stringList2 = new List<string>();
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取第一行的值列表
            for (int i = 0; i <= sheet.Cells.MaxDataColumn; i++)
            {
                Cell cell = sheet.Cells[0, i];
                if (cell != null)
                {
                    string value = cell.StringValue;
                    if (!stringList1.Contains(value) && value != "")
                    {
                        stringList1.Add(value);
                    }
                }
            }
            // 获取第二行的值列表
            for (int i = 0; i <= sheet.Cells.MaxDataColumn; i++)
            {
                Cell cell = sheet.Cells[1, i];
                if (cell != null)
                {
                    string value = cell.StringValue;
                    if (!stringList2.Contains(value) && value != "")
                    {
                        stringList2.Add(value);
                    }
                }
            }
            wb.Dispose();

            // 取列表最长的
            if (stringList2.Count > stringList1.Count)
            {
                result = stringList2;
            }
            else
            {
                result = stringList1;
            }

            // 返回result
            return result;
        }

        // 从Excel文件中获取值列表
        public static List<string> GetColValueFromExcel(string excelPath, string colField)
        {
            // 定义列表
            List<string> result = new List<string>();
            List<string> stringList1 = new List<string>();
            List<string> stringList2 = new List<string>();
            int rowIndex = -1;
            int colIndex = -1;
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取第一行的值列表
            for (int i = 0; i <= sheet.Cells.MaxDataColumn; i++)
            {
                Cell cell = sheet.Cells[0, i];
                if (cell != null)
                {
                    string value = cell.StringValue;
                    if (!stringList1.Contains(value) && value != "")
                    {
                        stringList1.Add(value);
                    }
                }
            }
            // 获取第二行的值列表
            for (int i = 0; i <= sheet.Cells.MaxDataColumn; i++)
            {
                Cell cell = sheet.Cells[1, i];
                if (cell != null)
                {
                    string value = cell.StringValue;
                    if (!stringList2.Contains(value) && value != "")
                    {
                        stringList2.Add(value);
                    }
                }
            }

            // 取列表最长的
            if (stringList2.Count > stringList1.Count)
            {
                rowIndex = 1;
            }
            else
            {
                rowIndex = 0;
            }

            // 获取表头所在的列
            for (int i = 0; i <= sheet.Cells.MaxDataColumn; i++)
            {
                Cell cell = sheet.Cells[rowIndex, i];
                if (cell != null)
                {
                    string value = cell.StringValue;
                    if (value == colField)
                    {
                        colIndex = i;
                        break;
                    }
                }
            }

            // 获取表头所在的列的值列表
            for (int i = 0; i <= sheet.Cells.MaxDataRow; i++)
            {
                Cell cell = sheet.Cells[i, colIndex];
                if (cell != null)
                {
                    string value = cell.StringValue;
                    if (!result.Contains(value) && value != "" && value != colField)
                    {
                        result.Add(value);
                    }
                }
            }

            wb.Dispose();
            // 返回result
            return result;
        }

        // 从Excel文件中获取Dictionary
        public static Dictionary<string, string> GetDictFromExcel(string excelPath, int col1 = 0, int col2 = 1)
        {
            // 定义字典
            Dictionary<string, string> dict = new Dictionary<string, string>();
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取key和value值
            for (int i = 0; i <= sheet.Cells.MaxDataRow; i++)
            {
                Cell key = sheet.Cells[i, col1];
                Cell value = sheet.Cells[i, col2];
                if (key != null && value != null)
                {
                    if (!dict.ContainsKey(key.StringValue))
                    {
                        if (key.StringValue != "" && value.StringValue != "")   // 空值不纳入
                        {
                            dict.Add(key.StringValue, value.StringValue);
                        }
                    }
                }
            }
            wb.Dispose();
            // 返回dict
            return dict;
        }

        // 从Excel文件中获取Dictionary列表
        public static List<Dictionary<string, string>> GetDictListFromExcelCol(string excelPath)
        {
            // 定义字典
            List<Dictionary<string, string>> list = new List<Dictionary<string, string>>();
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];

            for (int k = 1; k <= sheet.Cells.MaxDataRow; k++)
            {
                Dictionary<string, string> dict = new Dictionary<string, string>();
                for (int i = 0; i <= sheet.Cells.MaxDataColumn; i++)
                {
                    Cell key = sheet.Cells[0, i];
                    Cell value = sheet.Cells[k, i];
                    if (key != null && value != null)
                    {
                        if (!dict.ContainsKey(key.StringValue))
                        {
                            if (key.StringValue != "" && value.StringValue != "")   // 空值不纳入
                            {
                                dict.Add(key.StringValue, value.StringValue);
                            }
                        }
                    }
                }
                list.Add(dict);
            }

            wb.Dispose();
            // 返回dict
            return list;
        }

        // 从Excel文件中获取Dictionary
        public static Dictionary<string, double> GetDictFromExcelDouble(string excelPath, int col1 = 0, int col2 = 1)
        {
            // 定义字典
            Dictionary<string, double> dict = new Dictionary<string, double>();
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取key和value值
            for (int i = 0; i <= sheet.Cells.MaxDataRow; i++)
            {
                Cell key = sheet.Cells[i, col1];
                Cell value = sheet.Cells[i, col2];
                if (key != null && value != null)
                {
                    if (!dict.ContainsKey(key.StringValue))
                    {
                        if (key.StringValue != "" && value.StringValue != "")   // 空值不纳入
                        {
                            dict.Add(key.StringValue, value.DoubleValue);
                        }
                    }
                }
            }
            wb.Dispose();
            // 返回dict
            return dict;
        }

        // 从Excel文件中获取Dictionary
        public static Dictionary<string, string> GetDictFromExcelAll(string excelPath, int col1 = 0, int col2 = 1)
        {
            // 定义字典
            Dictionary<string, string> dict = new Dictionary<string, string>();
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取key和value值
            for (int i = 0; i <= sheet.Cells.MaxDataRow; i++)
            {
                Cell key = sheet.Cells[i, col1];
                Cell value = sheet.Cells[i, col2];
                if (key != null && value != null)
                {
                    if (!dict.ContainsKey(key.StringValue))
                    {
                        dict.Add(key.StringValue, value.StringValue);
                    }
                }
            }
            wb.Dispose();
            // 返回dict
            return dict;
        }

        // 从Excel文件中获取List
        public static List<string> GetListFromExcel(string excelPath, int col, int startRow = 0)
        {
            // 定义列表
            List<string> list = new List<string>();
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取key和value值
            for (int i = startRow; i <= sheet.Cells.MaxDataRow; i++)
            {
                Cell cell = sheet.Cells[i, col];
                if (cell != null)
                {
                    string strValue = cell.StringValue;
                    if (!list.Contains(strValue) && strValue != "")
                    {
                        list.Add(strValue);
                    }
                }
            }
            wb.Dispose();
            // 返回list
            return list;
        }

        // 从Excel文件中获取List
        public static List<string> GetListFromExcelAll(string excelPath, int col, int startRow = 0)
        {
            // 定义列表
            List<string> list = new List<string>();
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取key和value值
            for (int i = startRow; i <= sheet.Cells.MaxDataRow; i++)
            {
                Cell cell = sheet.Cells[i, col];
                if (cell != null)
                {
                    string strValue = cell.StringValue;
                    list.Add(strValue);
                }
            }
            wb.Dispose();
            // 返回list
            return list;
        }

        // 从Excel文件中获取Cellvalue
        public static string GetCellFromExcel(string excelPath, int row, int col)
        {
            // 定义value
            string value = "";
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取value值
            Cell cell = sheet.Cells[row, col];
            if (cell != null)
            {
                value = cell.StringValue;
            }
            wb.Dispose();
            // 返回value
            return value;
        }

        //  插入行
        public static void InsertRows(string excelPath, int rowIndex, int totalRows)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);

            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];

            Cells cells = sheet.Cells;

            // 插入行
            cells.InsertRows(rowIndex, totalRows);
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel指定范围导出JPG图片
        public static void ImportToJPG(string excelPath, string outputPath)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 设置打印属性
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            // 在一页内打印
            imgOptions.OnePagePerSheet = true;
            // 只打印区域内
            imgOptions.OnlyArea = true;
            // 打印
            SheetRender render = new SheetRender(sheet, imgOptions);
            render.ToImage(0, outputPath);
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel导出PDF
        public static void ImportToPDF(string excelPath, string pdfPath)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 设置PDF保存选项（可选）
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // 将整个Excel文件导出为PDF，所有工作表会被合并到一个PDF中
            wb.Save(pdfPath, SaveFormat.Pdf);
        }

        // Excel文件导出图片(阿来来)
        public static void Sheet2Pic(string excelPath)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];

            sheet.PageSetup.LeftMargin = 1;
            sheet.PageSetup.RightMargin = 1;
            sheet.PageSetup.BottomMargin = 1;
            sheet.PageSetup.TopMargin = 1;

            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            imgOptions.OnePagePerSheet = true;
            imgOptions.PrintingPage = PrintingPageType.IgnoreBlank;

            SheetRender sr = new SheetRender(sheet, imgOptions);
            string parentDirectory = System.IO.Directory.GetParent(excelPath).FullName;
            string wbName = wb.FileName[(wb.FileName.LastIndexOf(@"\") + 1)..].Replace(".xslx", "");

            string pathsave = parentDirectory + $@"\{wbName}.jpg";
            sr.ToImage(0, pathsave);

            // 保存
            wb.Dispose();
        }

        // 合并Excel文件同值列
        public static void MergeSameCol(string excelPath, int sheet_in_col, int startRow = 0)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            Cells cells = sheet.Cells;

            // 初始化列表
            string previous = "";
            List<List<int>> rowsList = new List<List<int>>();
            List<int> newRows = new List<int>();
            // 逐行处理
            for (int i = startRow; i <= cells.MaxDataRow + 1; i++)
            {
                //  获取目标cell
                Cell inCell = cells[i, sheet_in_col];

                // 属性映射
                if (inCell is not null)
                {
                    string va = inCell.StringValue;
                    if (va != previous)
                    {
                        // 如果不是刚开始的时候，并且合并行大于1时，就把列表发达给rowsList
                        if (previous != "" && newRows[1] > 1)
                        {
                            rowsList.Add(new List<int> { newRows[0], newRows[1] });
                            previous = va;
                            newRows[0] = i;     // 起始格
                            newRows[1] = 1;    // 合并行数
                        }
                        else
                        {
                            previous = va;
                            newRows.Clear();   // 清理掉原来的
                            newRows.Add(i);   // 起始格
                            newRows.Add(1);   // 合并行数
                        }
                    }
                    else
                    {
                        newRows[1] += 1;    // 合并行数+1
                    }
                }
            }

            // 合并处理
            foreach (var rows in rowsList)
            {
                int start = rows[0];
                int count = rows[1];
                try
                {
                    cells.Merge(start, sheet_in_col, count, 1);
                }
                catch (Exception)
                {

                    continue;
                }

            }

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // 合并Excel单元格
        public static void MergeCell(string excelPath, int start_row, int start_col, int total_rows, int total_cols)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            Cells cells = sheet.Cells;
            // 合并处理
            cells.Merge(start_row, start_col, total_rows, total_cols);

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // 合并Excel文件同值行
        public static void MergeSameRow(string excelPath, int sheet_in_row, int startCol = 0, int endCol = 0)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            Cells cells = sheet.Cells;

            // 初始化列表
            string previous = "";
            List<List<int>> colsList = new List<List<int>>();
            List<int> newCols = new List<int>();
            // 分析处理的列数
            int colCount = 0;
            if (endCol == 0)
            {
                colCount = cells.MaxDataColumn;
            }
            else
            {
                colCount = endCol + 1;
            }

            // 逐行处理
            for (int i = startCol; i < colCount; i++)
            {
                //  获取目标cell
                Cell inCell = cells[sheet_in_row, i];

                // 属性映射
                if (inCell is not null)
                {
                    string va = inCell.StringValue;
                    if (va != previous)
                    {
                        // 如果不是刚开始的时候，并且合并行大于1时，就把列表发达给rowsList
                        if (previous != "" && newCols[1] > 1)
                        {
                            colsList.Add(new List<int> { newCols[0], newCols[1] });
                            previous = va;
                            newCols[0] = i;     // 起始格
                            newCols[1] = 1;    // 合并行数
                        }
                        else
                        {
                            previous = va;
                            newCols.Clear();   // 清理掉原来的
                            newCols.Add(i);   // 起始格
                            newCols.Add(1);   // 合并行数
                        }
                    }
                    else
                    {
                        newCols[1] += 1;    // 合并行数+1
                    }
                }
            }

            // 合并处理
            foreach (var cols in colsList)
            {
                int start = cols[0];
                int count = cols[1];
                try
                {
                    cells.Merge(sheet_in_row, start, count, 1);
                }
                catch (Exception)
                {
                    continue;
                }

            }

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // 分割Excel
        public static void SpliteSheets(string excelPath, string folderPath)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            //  获取工作表
            WorksheetCollection sheets = wb.Worksheets;
            // 分割
            foreach (Worksheet sheet in sheets)
            {
                // 筛选一下，如果sheet没内容，就跳过
                Cells cells = sheet.Cells;
                if (cells.MaxDataRow < 0)
                {
                    continue;
                }
                // 新建工作薄
                Workbook newWb = new Workbook();
                newWb.Worksheets[0].Copy(sheet);

                // 保存
                newWb.Save(@$"{folderPath}\{sheet.Name}.xlsx", SaveFormat.Xlsx);
                newWb.Dispose();
            }
            wb.Dispose();
        }

        // 合并Excel
        public static void MergeSheets(string excelPath, string folderPath, bool isExcel)
        {
            // 获取所有excel文件
            List<string> xlsList = DirTool.GetAllFiles(folderPath, ".xls");
            List<string> xlsxList = DirTool.GetAllFiles(folderPath, ".xlsx");
            xlsList.AddRange(xlsxList);

            // 复制Excel表
            DirTool.CopyResourceFile(@"CCTool.Data.Excel.界线描述表.xlsx", excelPath);

            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);

            // 打开工作薄
            Workbook targetWB = OpenWorkbook(excelFile);

            // 合并
            foreach (string xlsFile in xlsList)
            {
                // 打开工作薄
                Workbook originWB = OpenWorkbook(xlsFile);
                //  获取工作表
                WorksheetCollection sheets = originWB.Worksheets;
                foreach (Worksheet sheet in sheets)
                {
                    // 筛选一下，如果sheet没内容，就跳过
                    Cells cells = sheet.Cells;
                    if (cells.MaxDataRow < 0)
                    {
                        continue;
                    }
                    // 获取sheet名称
                    string sheetName = "范例";
                    string excelName = xlsFile[(xlsFile.LastIndexOf(@"\") + 1)..xlsFile.LastIndexOf(@".")];
                    if (isExcel)
                    {
                        sheetName = excelName;
                    }
                    else
                    {
                        sheetName = $"{excelName}_{sheet.Name}";
                    }
                    // 插入一个空白页
                    targetWB.Worksheets.Add(sheetName);
                    targetWB.Worksheets[sheetName].Copy(sheet);
                    // 删除第一页
                    targetWB.Worksheets.RemoveAt("sheet1");
                }
                originWB.Dispose();
            }
            // 保存
            targetWB.Save(excelPath);
            targetWB.Dispose();
        }

        // 设置表格单元为数字型及保留位数
        public static void SetDigit(string excelPath, List<int> cols, int startRow, int digit)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            Cells cells = sheet.Cells;

            // 设置单元格为数字型，小数位数
            foreach (int col in cols)
            {
                for (int i = startRow; i <= cells.MaxDataRow; i++)
                {
                    Style style = sheet.Cells[i, col].GetStyle();
                    style.Number = 4;   // 数字型
                    style.Custom = digit switch   // 小数位数
                    {
                        1 => "0.0",
                        2 => "0.00",
                        3 => "0.000",
                        4 => "0.0000",
                        5 => "0.00000",
                        6 => "0.000000",
                        _ => null,
                    };

                    sheet.Cells[i, col].SetStyle(style);
                }
            }

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // Excel导出HTML
        public static void Sheet2Html(string excelPath, string htmlPath)
        {
            // 获取工作薄、工作表
            string excelFile = GetPath(excelPath);
            int sheetIndex = GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = OpenWorkbook(excelFile);
            // 设置为当前活动的表
            wb.Worksheets.ActiveSheetIndex = sheetIndex;

            // 创建保存选项
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // 如果需要导出单个工作表而不是整个工作簿，可以设置 ExportActiveWorksheetOnly 为 true
                ExportActiveWorksheetOnly = true,

                // 设置 HTML 文件中包含的图片和样式信息
                ExportImagesAsBase64 = true
            };

            // 导出为 HTML 文件
            wb.Save(htmlPath, saveOptions);
        }

    }
}
