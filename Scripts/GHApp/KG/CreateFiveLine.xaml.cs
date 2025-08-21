using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
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

namespace CCTool.Scripts.GHApp.KG
{
    /// <summary>
    /// Interaction logic for CreateFiveLine.xaml
    /// </summary>
    public partial class CreateFiveLine : ArcGIS.Desktop.Framework.Controls.ProWindow
    {

        // 工具设置标签
        readonly string toolSet = "CreateFiveLine";

        public CreateFiveLine()
        {
            InitializeComponent();

            try
            {
                // 初始化参数选项
                textGDBPath.Text = BaseTool.ReadValueFromReg(toolSet, "gdbPath");
                textFront.Text = BaseTool.ReadValueFromReg(toolSet, "front");
                cb_addToMap.IsChecked = BaseTool.ReadValueFromReg(toolSet, "addToMap").ToBool();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

            
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "生成五线";

        private void openGDBButton_Click(object sender, RoutedEventArgs e)
        {
            textGDBPath.Text = UITool.OpenDialogGDB();
        }

        // 添加要素图层的所有字符串字段到combox中
        private void combox_field_DropDown(object sender, EventArgs e)
        {
            // 将图层字段加入到Combox列表中
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_bmField);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147592512?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        // 运行
        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                var defPath = Project.Current.HomeFolderPath;
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string field_bm = combox_bmField.ComboxText();

                string gdbPath = textGDBPath.Text;
                string front = textFront.Text;

                bool addToMap = (bool)cb_addToMap.IsChecked;

                // 判断参数是否选择完全
                if (fc_path == "" || field_bm == "" || gdbPath == "" || front == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 保存参数
                BaseTool.WriteValueToReg(toolSet, "gdbPath", gdbPath);
                BaseTool.WriteValueToReg(toolSet, "front", front);
                BaseTool.WriteValueToReg(toolSet, "addToMap", addToMap);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(fc_path, field_bm, front);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10, "提取红线");
                    // 提取红线
                    string hx = $@"{gdbPath}\{front}_红线";
                    string sql_hx = $@"{field_bm} IN ('1207')";
                    Arcpy.Select(fc_path, hx, sql_hx);
                    // 统计图斑数
                    pw.AddMessageMiddle(5, $"    红线的图斑数为：【{hx.GetFeatureCount()}】", Brushes.Gray);

                    pw.AddMessageMiddle(10, "提取绿线");
                    // 提取绿线
                    string lvx = $@"{gdbPath}\{front}_绿线";
                    string sql_lvx = $@"{field_bm} IN ('1401', '1402', '1403')";
                    Arcpy.Select(fc_path, lvx, sql_lvx);
                    // 统计图斑数
                    pw.AddMessageMiddle(5, $"    绿线的图斑数为：【{lvx.GetFeatureCount()}】", Brushes.Gray);


                    pw.AddMessageMiddle(10, "提取蓝线");
                    // 提取蓝线
                    string lx = $@"{gdbPath}\{front}_蓝线";
                    string sql_lx = $@"{field_bm} IN ('1701', '1702', '1703', '1704', '1705', '1706')";
                    Arcpy.Select(fc_path, lx, sql_lx);
                    // 统计图斑数
                    pw.AddMessageMiddle(5, $"    蓝线的图斑数为：【{lx.GetFeatureCount()}】", Brushes.Gray);

                    pw.AddMessageMiddle(10, "提取紫线");
                    // 提取紫线
                    string zx = $@"{gdbPath}\{front}_紫线";
                    string sql_zx = $@"{field_bm} IN ('1504')";
                    Arcpy.Select(fc_path, zx, sql_zx);
                    // 统计图斑数
                    pw.AddMessageMiddle(5, $"    紫线的图斑数为：【{zx.GetFeatureCount()}】", Brushes.Gray);

                    pw.AddMessageMiddle(10, "提取黄线");
                    // 提取黄线
                    string hux = $@"{gdbPath}\{front}_黄线";
                    string sql_hux = $@"{field_bm} IN ('1208', '120801', '120802', '120803', '1203', '1206', '1204', '1301', '1302', '1303', '1304', '1305', '1306', '1307', '1308', '1309')";
                    Arcpy.Select(fc_path, hux, sql_hux);
                    // 统计图斑数
                    pw.AddMessageMiddle(5, $"    黄线的图斑数为：【{hux.GetFeatureCount()}】", Brushes.Gray);

                    if (addToMap)
                    {
                        pw.AddMessageMiddle(5, $"加载五线");
                        // 加载图层
                        List<string> lines = new List<string>() { "红线", "绿线", "蓝线", "紫线", "黄线" };
                        foreach (string line in lines)
                        {
                            string path = $@"{gdbPath}\{front}_{line}";
                            long count = path.GetFeatureCount();
                            if (count > 0)
                            {
                                MapCtlTool.AddLayerToMap(path);
                                string lyPath = $@"{defPath}\WX_{line}.lyrx";
                                DirTool.CopyResourceFile($@"CCTool.Data.Layers.WX_{line}.lyrx", lyPath);
                                // 应用图层符号
                                string lyName = $@"{front}_{line}";
                                FeatureLayer featureLayer = lyName.TargetFeatureLayer();
                                GisTool.ApplySymbol(featureLayer, lyPath);

                                File.Delete(lyPath);
                            }
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

        private List<string> CheckData(string fc, string field, string front)
        {
            List<string> result = new List<string>();

            // 检查字段值是否符合要求
            string result_value = CheckTool.CheckFieldValue(fc, field, GlobalData.dic_ydyh_new.Keys.ToList());
            if (result_value != "")
            {
                result.Add(result_value);
            }

            // 检查前缀是否是数字
            if (front.IsNumeric())
            {
                result.Add($"前缀不能以数字开头。\r");
            }

            return result;
        }
    }
}
