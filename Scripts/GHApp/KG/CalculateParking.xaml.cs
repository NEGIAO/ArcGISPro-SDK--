using ActiproSoftware.Products.Ribbon;
using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using Aspose.Words.Fields;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
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

namespace CCTool.Scripts.GHApp.KG
{
    /// <summary>
    /// Interaction logic for CalculateParking.xaml
    /// </summary>
    public partial class CalculateParking : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "CalculateParking";

        public CalculateParking()
        {
            InitializeComponent();

            txtMJ.Text = BaseTool.ReadValueFromReg(toolSet, "ppArea");

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "计算停车位(福建标准)";


        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147729900?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        // 运行
        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                var defGDB = Project.Current.DefaultGeodatabasePath;
                // 获取参数
                string yd = combox_yd.ComboxText();
                string bmField = combox_bmField.ComboxText();
                string mjField = combox_mjField.ComboxText();
                string parkField = combox_parkField.ComboxText();
                string farField = combox_farField.ComboxText();
                string ssField = combox_ss.ComboxText();
                double ppArea = txtMJ.Text.ToDouble();    // 户均面积

                // 判断参数是否选择完全
                if (yd == "" || bmField == "" || mjField == "" || parkField == "" || farField == "" || ssField == "" || ppArea < 0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                BaseTool.WriteValueToReg(toolSet, "ppArea", ppArea);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(yd, bmField, mjField, parkField, farField);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(20, "按用地分类计算停车位");
                    Table table = yd.TargetTable();
                    using RowCursor rowCursor = table.Search();
                    while (rowCursor.MoveNext())
                    {
                        Row row = rowCursor.Current;

                        int parkCount = 0;

                        // 获取字段值
                        string bm = row[bmField]?.ToString() ?? "";     // 用地编码
                        string mjStr = row[mjField]?.ToString() ?? "";
                        string farStr = row[farField]?.ToString() ?? "";
                        string ss = row[ssField]?.ToString() ?? "";

                        double mj = mjStr.ToDouble();    // 用地面积
                        double far = farStr.ToDouble();   // 容积率

                        double jzmj = mj * far;    // 建筑面积

                        // 计算停车位
                        if (bm == "070101" || bm == "070102" || bm == "070103" || bm == "0701")
                        {
                            parkCount += (jzmj / ppArea * 1.2).ToString().ToInt();
                        }
                        //else if (bm == "0703")
                        //{
                        //    parkCount += (jzmj / ppArea * 0.3).ToString().ToInt();
                        //}
                        else if (bm == "0901" || bm == "090101")
                        {
                            parkCount += (jzmj / 100 * 0.6).ToString().ToInt();
                        }
                        else if (bm == "090102" || bm == "090103" || bm == "090301" || bm == "0902")
                        {
                            parkCount += (jzmj / 100 * 1.2).ToString().ToInt();
                        }
                        else if (bm == "090104")
                        {
                            parkCount += (jzmj / 50 * 0.3).ToString().ToInt();
                        }
                        else if (bm == "0806" || bm == "0801" || bm == "0803" || bm == "0807")
                        {
                            parkCount += (jzmj / 100 * 0.8).ToString().ToInt();
                        }
                        // 停车场
                        else if (bm == "120803")
                        {
                            parkCount += (mj / 25).ToString().ToInt();
                        }
                        // 幼儿园
                        else if (bm == "080404" && (ss.Contains("幼儿园") || ss.Contains("幼托")))
                        {
                            string a = (mj / 15 / 100 * 1.5).ToString();
                            int b = a.ToInt();
                            parkCount += b;
                        }
                        // 小学
                        else if (bm == "080403" && ss.Contains("小学"))
                        {
                            parkCount += (mj / 15 / 100 * 2).ToString().ToInt();
                        }
                        // 中学
                        else if (bm == "080403" && ss.Contains("中学"))
                        {
                            parkCount += (mj / 18 / 100 * 3).ToString().ToInt();
                        }

                        row[parkField] = parkCount;
                        row.Store();
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

        private List<string> CheckData(string yd, string bmField, string mjField, string parkField, string farField)
        {
            List<string> result = new List<string>();

            List<string> fields = new List<string>() { bmField, mjField, parkField, farField };
            // 检查字段是否存在
            string result_value = CheckTool.IsHaveFieldInTarget(yd, fields);
            if (result_value != "")
            {
                result.Add(result_value);
            }

            return result;
        }

        private void combox_yd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_yd);
        }

        private void combox_bmField_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_yd.ComboxText(), combox_bmField);

        }

        private void combox_mjField_DropDown(object sender, EventArgs e)
        {
            UITool.AddAllFloatFieldsToComboxPlus(combox_yd.ComboxText(), combox_mjField);
        }

        private void combox_parkField_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_yd.ComboxText(), combox_parkField);
        }

        private void combox_farField_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_yd.ComboxText(), combox_farField);
        }

        private void combox_ss_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_yd.ComboxText(), combox_ss);
        }

        private async void combox_yd_DropClose(object sender, EventArgs e)
        {
            try
            {
                string yd = combox_yd.ComboxText();

                // 初始化参数选项
                await UITool.InitLayerFieldToComboxPlus(combox_bmField, yd, "YDFLDM", "string");
                await UITool.InitLayerFieldToComboxPlus(combox_mjField, yd, "Shape_Area", "float");
                await UITool.InitLayerFieldToComboxPlus(combox_farField, yd, "RJLSX", "float");
                await UITool.InitLayerFieldToComboxPlus(combox_ss, yd, "PTSS", "string");
                await UITool.InitLayerFieldToComboxPlus(combox_parkField, yd, "PTTCBW", "string");

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
            
        }
    }
}
