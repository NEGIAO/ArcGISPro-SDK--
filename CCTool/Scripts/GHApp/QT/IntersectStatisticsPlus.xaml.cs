using ArcGIS.Core.Data;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Mapping.Symbology;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using SharpCompress.Common;
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

namespace CCTool.Scripts.GHApp.QT
{
    /// <summary>
    /// Interaction logic for IntersectStatisticsPlus.xaml
    /// </summary>
    public partial class IntersectStatisticsPlus : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public IntersectStatisticsPlus()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "林地占比分析强化版";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                var def_gdb = Project.Current.DefaultGeodatabasePath;
                // 获取要素
                string slzy = combox_origin.ComboxText();
                string ld = combox_identy.ComboxText();

                // 判断参数是否选择完全
                if (slzy == "" || ld == "")
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
                    List<string> lines = new List<string>() { slzy, ld };
                    // 检查数据
                    List<string> errs = CheckData(lines);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(0, err, Brushes.Red);
                        }
                        return;
                    }

                    // 计算主导功能
                    pw.AddMessageMiddle(20, "计算主导功能");
                    StatisticsField(slzy, ld, "ZDGN", "ZDGN", pw);

                    // 计算林种
                    pw.AddMessageMiddle(20, "计算林种");
                    StatisticsField(slzy, ld, "LZJS", "LZ", pw);

                    // 计算蓄积量
                    pw.AddMessageMiddle(20, "计算蓄积量");
                    StatisticsXJ(slzy, ld, "XJ率", "ZXJL", pw);


                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        // 汇总统计字段指标
        private void StatisticsField(string slzy, string ld, string field, string targetField, ProcessWindow pw)
        {
            // 获取默认数据库
            var def_gdb = Project.Current.DefaultGeodatabasePath;

            pw.AddMessageMiddle(0, "汇总并整理", Brushes.Gray);
            string bjField = "BKH";
            // 标识
            string slzy_copy = $@"{def_gdb}\slzy_copy";
            string identity = $@"{def_gdb}\identity";
            Arcpy.CopyFeatures(slzy, slzy_copy);
            Arcpy.DeleteField(slzy_copy, bjField, "KEEP_FIELDS");
            Arcpy.Identity(ld, slzy_copy, identity);

            // 汇总并整理
            string statistics = $@"{def_gdb}\statistics";
            string table_sta = $@"{def_gdb}\table_sta";
            string table_sort = $@"{def_gdb}\table_sort";
            Arcpy.Statistics(identity, statistics, $"shape_area SUM", $"{bjField};{field}");
            Arcpy.TableSelect(statistics, table_sta, $"{bjField} <> ''");
            Arcpy.Sort(table_sta, table_sort, $"{bjField} ASCENDING;{field} DESCENDING", "UR");
            Arcpy.CalculateField(table_sort, field, $"ss(!{field}!)", "def ss(a):\r\n    if a == None:\r\n        return \"待确认\"\r\n    elif a.replace(\" \",\"\") ==\"\":\r\n        return \"待确认\"\r\n    else:\r\n        return a");

            pw.AddMessageMiddle(0, "提取指标", Brushes.Gray);
            // 提取指标
            var dic_slzy = GisTool.GetDictFromPathDouble(slzy_copy, bjField, "shape_area");
            Dictionary<string, string> dic = new();

            Table table = table_sort.TargetTable();
            using RowCursor rowCursor = table.Search();
            while (rowCursor.MoveNext())
            {
                using Row row = rowCursor.Current;
                // 获取value
                var bj = row[bjField];
                var fd = row[field];
                var area = row["SUM_SHAPE_Area"];
                if (bj != null && area != null)
                {
                    string bjValue = bj.ToString();
                    string fdValue = fd.ToString();
                    double areaValue = double.Parse(area.ToString());
                    // 对应图斑的总面积和比例
                    double totalArea = dic_slzy[bjValue];
                    double preValue = Math.Round(areaValue / totalArea * 100, 1);
                    // 输出文字
                    string result = $"{fdValue}{preValue}%，";
                    // 返回
                    if (dic.ContainsKey(bjValue))
                    {
                        dic[bjValue] += result;
                    }
                    else
                    {
                        dic[bjValue] = result;
                    }
                }
            }


            pw.AddMessageMiddle(0, "指标赋值", Brushes.Gray);
            // 指标赋值给森林资源斑块
            Table table_slzy = slzy.TargetTable();
            using RowCursor rowCursor2 = table_slzy.Search();
            while (rowCursor2.MoveNext())
            {
                using Row row = rowCursor2.Current;
                // 获取value
                var bj = row[bjField];
                if (bj != null)
                {
                    string bjValue = bj.ToString();
                    row[targetField] = "";
                    if (dic.ContainsKey(bjValue))
                    {
                        string va = dic[bjValue].ToString();
                        row[targetField] = va[..(va.Length - 1)];
                    }
                }
                row.Store();
            }
        }

        // 计算蓄积量
        private void StatisticsXJ(string slzy, string ld, string field, string targetField, ProcessWindow pw)
        {
            // 获取默认数据库
            var def_gdb = Project.Current.DefaultGeodatabasePath;

            pw.AddMessageMiddle(0, "汇总并整理", Brushes.Gray);
            string bjField = "BKH";
            // 标识
            string slzy_copy = $@"{def_gdb}\slzy_copy";
            string identity = $@"{def_gdb}\identity";
            Arcpy.CopyFeatures(slzy, slzy_copy);
            Arcpy.DeleteField(slzy_copy, bjField, "KEEP_FIELDS");
            Arcpy.Identity(ld, slzy_copy, identity);

            // 汇总并整理
            Arcpy.AddField(identity, "MJJ", "DOUBLE");
            Arcpy.CalculateField(identity, "MJJ", $"ss(!SHAPE_Area!,!{field}!)", "def ss(a,b):\r\n    if b is None:\r\n        return 0\r\n    else:\r\n        return round(a*b, 1)");
            string statistics = $@"{def_gdb}\statistics";
            string table_sta = $@"{def_gdb}\table_sta";
            string table_sort = $@"{def_gdb}\table_sort";
            Arcpy.Statistics(identity, statistics, $"MJJ SUM", $"{bjField};{field}");
            Arcpy.TableSelect(statistics, table_sta, $"{bjField} <> ''");
            Arcpy.Sort(table_sta, table_sort, $"{bjField} ASCENDING;{field} DESCENDING", "UR");

            pw.AddMessageMiddle(0, "提取指标", Brushes.Gray);
            // 提取指标
            Dictionary<string, double> dic = new();

            Table table = table_sort.TargetTable();
            using RowCursor rowCursor = table.Search();
            while (rowCursor.MoveNext())
            {
                using Row row = rowCursor.Current;
                // 获取value
                var bj = row[bjField];
                var area = row["SUM_MJJ"];
                if (bj != null && area != null)
                {
                    string bjValue = bj.ToString();
                    // 蓄积量
                    double xjl = double.Parse(area.ToString());
                    
                    // 返回
                    if (xjl != 0)
                    {
                        if (dic.ContainsKey(bjValue))
                        {
                            dic[bjValue] += xjl;
                        }
                        else
                        {
                            dic[bjValue] = xjl;
                        }
                    }
                }
            }


            pw.AddMessageMiddle(0, "指标赋值", Brushes.Gray);
            // 指标赋值给森林资源斑块
            Table table_slzy = slzy.TargetTable();
            using RowCursor rowCursor2 = table_slzy.Search();
            while (rowCursor2.MoveNext())
            {
                using Row row = rowCursor2.Current;
                // 获取value
                var bj = row[bjField];
                if (bj != null)
                {
                    string bjValue = bj.ToString();
                    row[targetField] = 0;
                    if (dic.ContainsKey(bjValue))
                    {
                        row[targetField] = dic[bjValue];
                    }
                }
                row.Store();
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/140240696";
            UITool.Link2Web(url);
        }


        private void combox_origin_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_origin);
        }


        private void combox_identy_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_identy);
        }


        private List<string> CheckData(List<string> lines)
        {
            List<string> result = new List<string>();


            return result;
        }

    }
}
