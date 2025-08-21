using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Shared;
using NPOI.SS.Formula.Functions;
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

namespace CCTool.Scripts.GDBMenu
{
    /// <summary>
    /// Interaction logic for CalculateYSDM.xaml
    /// </summary>
    public partial class CalculateYSDM : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public CalculateYSDM()
        {
            InitializeComponent();

            // 读取excel
            textExcelPath.Text = BaseTool.ReadValueFromReg("YSDMTool", "initPath");
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            string path = UITool.OpenDialogExcel();
            textExcelPath.Text = path;

            BaseTool.WriteValueToReg("YSDMTool", "initPath", path);
        }

        private void downLoadButton_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://pan.baidu.com/s/1J1eCn1XNKZ5I2gEuRMOJ8w?pwd=a8e9";
            UITool.Link2Web(url);
        }

        private void btn_help_click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/141822787?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "整库计算YSDM";

        private async void btn_go_click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取数据
                string excelPath = textExcelPath.Text;
                string fieldName = "YSDM";

                // 判断参数是否选择完全
                if (excelPath == "")
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
                    // 获取当前选择的gdb数据库
                    string gdbPath = Project.Current.SelectedItems.OfType<GDBProjectItem>().FirstOrDefault().Path;
                    // 要素类
                    List<string> originFCList = gdbPath.GetFeatureClassPathFromGDB();
                    // 独立表
                    List<string> tbList = gdbPath.GetStandaloneTablePathFromGDB();
                    // 合并
                    originFCList.AddRange(tbList);

                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(originFCList);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    // 获取excel列表
                    var pairs = ExcelTool.GetDictFromExcel(excelPath);

                    foreach (string fc in originFCList)
                    {
                        // 没有BSM的话，跳过
                        if (!GisTool.IsHaveFieldInTarget(fc, fieldName))
                        {
                            continue;
                        }

                        string fcName = fc[(fc.LastIndexOf(".gdb") + 5)..];
                        
                        //  获取要素的的名称或别名
                        string name = fc.GetFeatureClassAtt().Name;
                        string aliasName = fc.GetFeatureClassAtt().AliasName;

                        // 如果符合列表
                        if (pairs.ContainsKey(name))
                        {
                            string value = pairs[name];
                            pw.AddMessageMiddle(10, $"计算YSDM：{fcName}…………{value}");
                            // 计算BSM
                            Arcpy.CalculateField(fc, fieldName, $"'{value}'");
                        }
                        else if (pairs.ContainsKey(aliasName))
                        {
                            string value = pairs[aliasName];
                            pw.AddMessageMiddle(10, $"计算YSDM：{fcName}…………{value}");
                            // 计算BSM
                            Arcpy.CalculateField(fc, fieldName, $"'{value}'");
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


        private List<string> CheckData(List<string> originFCList)
        {
            List<string> result = new List<string>();
            string field = "YSDM";
            bool isHasTarget = false;

            foreach (string fc in originFCList)
            {
                // 检查是否有BSM字段
                bool isHave = GisTool.IsHaveFieldInTarget(fc, field);
                if (isHave)
                {
                    isHasTarget = true;
                }
            }

            if (!isHasTarget)
            {
                result.Add($"GDB中不包含符号条件的要素类或表(有YSDM字段)");
            }

            return result;
        }
    }
}
