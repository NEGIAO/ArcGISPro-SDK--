using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using System;
using ArcGIS.Core.Data;
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
using System.IO;
using CCTool.Scripts.ToolManagers;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.POIFS.Crypt.Dsig;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Library;

namespace CCTool.Scripts.UI.ProWindow
{
    /// <summary>
    /// Interaction logic for CheckYDChangePlus.xaml
    /// </summary>
    public partial class CheckYDChangePlus : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public CheckYDChangePlus()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "现状规划用地变化检查(村规)";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            // 获取指标
            string fc_xz_txt = combox_fc_xz.ComboxText();
            string fc_gh_txt = combox_fc_gh.ComboxText();
            string field_change = @"用地变化";

            // 判断参数是否选择完全
            if (fc_xz_txt == "" || fc_gh_txt == "")
            {
                MessageBox.Show("有必选参数为空！！！");
                return;
            }

            try
            {
                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                // 输出检查结果要素
                string def_path = Project.Current.HomeFolderPath;
                string DefalutGDB = Project.Current.DefaultGeodatabasePath;
                string identityFeatureClass = DefalutGDB + @"\identityFeatureClass";
                string checkRezult = DefalutGDB + @"\checkRezult";

                Close();
                // 异步执行
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(fc_xz_txt, fc_gh_txt);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    // 复制模板，并获取Dic
                    string excel_map = def_path + @"\用地用海_归纳建设非建设用地2.xlsx";
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.用地用海_归纳建设非建设用地2.xlsx", excel_map);
                    Dictionary<string, string> dic = ExcelTool.GetDictFromExcel(excel_map + @"\sheet1$");

                    pw.AddMessageMiddle(10, "处理输入字段");

                    // 复制输入要素
                    Arcpy.CopyFeatures(fc_xz_txt, DefalutGDB + @"\tem_xz");
                    Arcpy.CopyFeatures(fc_gh_txt, DefalutGDB + @"\tem_gh");
                    // 删除无关字段
                    List<string> xz_keep = new List<string>() { "JQDLBM", "JQDLMC", "CZCSXM" };
                    List<string> gh_keep = new List<string>() { "GHDLBM", "GHDLMC", "SSBJLX" };
                    Arcpy.DeleteField(DefalutGDB + @"\tem_xz", xz_keep, "KEEP_FIELDS");
                    Arcpy.DeleteField(DefalutGDB + @"\tem_gh", gh_keep, "KEEP_FIELDS");

                    pw.AddMessageMiddle(10, "标识要素");
                    // 标识
                    Arcpy.Identity(DefalutGDB + @"\tem_xz", DefalutGDB + @"\tem_gh", identityFeatureClass);

                    // 添加字段
                    Arcpy.AddField(identityFeatureClass, field_change, "TEXT");

                    pw.AddMessageMiddle(20, "计算字段，找出变化图斑");

                    // 计算字段，找出变化图斑
                    using Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(DefalutGDB)));
                    using FeatureClass featureClass = gdb.OpenDataset<FeatureClass>("identityFeatureClass");
                    using (RowCursor rowCursor = featureClass.Search(null, false))
                    {
                        while (rowCursor.MoveNext())
                        {
                            using Row row = rowCursor.Current;
                            // 获取检查字段的值
                            var fd_xz = row["JQDLBM"];
                            var fd_gh = row["GHDLBM"];
                            var CZCSXM = row["CZCSXM"];
                            var SSBJLX = row["SSBJLX"];
                            if (fd_xz is not null && fd_gh is not null)
                            {
                                if (fd_xz.ToString() != fd_gh.ToString())
                                {
                                    // 归纳建设用地属性
                                    string xz_js = dic[fd_xz.ToString()];
                                    string gh_js = dic[fd_gh.ToString()];
                                    // 判断城镇用地
                                    if (CZCSXM is not null)
                                    {
                                        if (CZCSXM.ToString().Contains("201") || CZCSXM.ToString().Contains("202"))
                                        {
                                            xz_js = "城镇用地";
                                        }
                                    }
                                    if (SSBJLX is not null)
                                    {
                                        if (SSBJLX.ToString().Contains("z") || CZCSXM.ToString().Contains("Z"))
                                        {
                                            gh_js = "城镇用地";
                                        }
                                    }
                                    // 赋值
                                    row[field_change] = @$"【{xz_js}】-->【{gh_js}】";
                                }
                                row.Store();
                            }
                        }
                    }
                    pw.AddMessageMiddle(40, "提取变化图斑");
                    // 提取变化图斑
                    string sql = "用地变化 IS NOT NULL";
                    Arcpy.Select(identityFeatureClass, checkRezult, sql, true);
                    // 删除过程要素
                    File.Delete(excel_map);
                    Arcpy.Delect(identityFeatureClass);
                    Arcpy.Delect(DefalutGDB + @"\tem_xz");
                    Arcpy.Delect(DefalutGDB + @"\tem_gh");
                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void combox_fc_xz_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc_xz);
        }

        private void combox_fc_gh_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc_gh);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135773697?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string xz_data, string gh_data)
        {
            List<string> result = new List<string>();

            if (xz_data != "")
            {
                // 检查是否存在字段
                List<string> xz_keep = new List<string>() { "JQDLBM", "JQDLMC", "CZCSXM" };
                string fieldRusult = CheckTool.IsHaveFieldInLayer(xz_data, xz_keep);
                if (fieldRusult != "")
                {
                    result.Add(fieldRusult);
                }
                // 检查JQDLBM是否合规
                string result_value = CheckTool.CheckFieldValue(xz_data, "JQDLBM", GlobalData.dic_ydyh_new.Keys.ToList());
                if (result_value != "")
                {
                    result.Add(result_value);
                }
            }

            if (gh_data != "")
            {
                // 检查是否存在字段
                List<string> gh_keep = new List<string>() { "GHDLBM", "GHDLMC", "SSBJLX" };
                string fieldRusult = CheckTool.IsHaveFieldInLayer(gh_data, gh_keep);
                if (fieldRusult != "")
                {
                    result.Add(fieldRusult);
                }
                // 检查GHDLBM是否合规
                string result_value = CheckTool.CheckFieldValue(gh_data, "GHDLBM", GlobalData.dic_ydyh_new.Keys.ToList());
                if (result_value != "")
                {
                    result.Add(result_value);
                }
            }

            if (xz_data != "" && gh_data != "")
            {
                // 检查现状规划用地是否完全重叠
                string outPath = Project.Current.DefaultGeodatabasePath + @"\symdiff";
                // 交集取反
                Arcpy.SymDiff(xz_data, gh_data, outPath);
                long count = outPath.TargetTable().GetCount();
                if (count > 0)
                {
                    result.Add("2个输入图层不完全重叠！");
                }
                // 删除
                Arcpy.Delect(outPath);
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
