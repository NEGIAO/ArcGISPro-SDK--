using ArcGIS.Core.CIM;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.POIFS.Crypt.Dsig;
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
    /// Interaction logic for CalculateBSM.xaml
    /// </summary>
    public partial class CalculateBSM : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public CalculateBSM()
        {
            InitializeComponent();

            // 读取前缀
            textBox_front.Text = BaseTool.ReadValueFromReg("BSMFrontText", "initText");
        }


        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "整库计算BSM";

        private async void btn_go_click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取数据
                string front = textBox_front.Text;
                bool isByLength = (bool)rb_bsmLength.IsChecked;

                string leng = textBox_len.Text ?? "";     // 自定义字段长度
                int BLength = int.Parse(leng);
                string fieldName = "BSM";

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    // 保存输入的前缀
                    BaseTool.WriteValueToReg("BSMFrontText", "initText", front);

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

                    foreach (string fc in originFCList)
                    {
                        // 没有BSM的话，跳过
                        if (!GisTool.IsHaveFieldInTarget(fc, fieldName))
                        {
                            continue;
                        }

                        string fcName = fc[(fc.LastIndexOf(".gdb") + 5)..];
                        pw.AddMessageMiddle(10, $"计算BSM：{fcName}");
                        // 获取OID字段
                        string oidField = fc.TargetIDFieldName();

                        // 字段长度
                        int len = BLength;
                        if (isByLength)
                        {
                            len = GisTool.GetFieldFromString(fc, fieldName).Length;
                        }

                        // 计算BSM
                        Arcpy.CalculateField(fc, fieldName, $"'{front}'+'0' * ({len} - len(str(!{oidField}!))-{front.Length}) + str(!{oidField}!)");
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

        private void btn_help_click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/141819926";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(List<string> originFCList)
        {
            List<string> result = new List<string>();
            string field = "BSM";
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
                result.Add($"GDB中不包含符号条件的要素类或表(有BSM字段)");
            }

            return result;
        }

    }

}
