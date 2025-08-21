using ArcGIS.Core.CIM;
using ArcGIS.Core.Data.Exceptions;
using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
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
using CCTool.Scripts.Manager;
using System.IO;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Windows;
using NPOI.OpenXmlFormats.Dml;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Extensions;

namespace CCTool.Scripts
{
    /// <summary>
    /// Interaction logic for YDYHChanger.xaml
    /// </summary>
    public partial class YDYHChanger : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public YDYHChanger()
        {
            InitializeComponent();
            // combox_model框中添加2种转换模式，默认【代码转名称】
            combox_model.Items.Add("代码转名称");
            combox_model.Items.Add("名称转代码");
            combox_model.SelectedIndex = 0;

            combox_version.Items.Add("旧版");
            combox_version.Items.Add("新版");
            combox_version.SelectedIndex = 1;
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "用地用海代码和名称转换";

        private void combox_field_before_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field_before);
        }

        private void combox_field_after_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field_after);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        // 执行
        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string field_before = combox_field_before.ComboxText();
                string field_after = combox_field_after.ComboxText();
                string model = combox_model.Text;
                string version = combox_version.Text;

                // 判断参数是否选择完全
                if (fc_path == "" || field_before == "" || field_after == "")
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
                    List<string> errs = CheckData(model, version, fc_path, field_before);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(0, err, Brushes.Red);
                        }
                        return;
                    }

                    string def_folder = Project.Current.HomeFolderPath;     // 工程默认文件夹位置

                    // 复制转换表
                    string excelName = "";
                    if (version == "旧版")
                    {
                        excelName = "用地用海_DM_to_MC";
                    }
                    else
                    {
                        excelName = "新版用地用海_DM_to_MC";
                    }
                    string output_excel = $@"{def_folder}\{excelName}.xlsx";
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.{excelName}.xlsx", output_excel);

                    // 转换模式
                    bool reserve = false;
                    if (model == "名称转代码") { reserve = true; }

                    pw.AddMessageMiddle(10, "开始转换...");
                    // 用地用海编码名称互转
                    ComboTool.AttributeMapper(fc_path, field_before, field_after, output_excel + @"\sheet1$", reserve);

                    pw.AddMessageMiddle(60, "删除中间数据");
                    // 删除中间数据
                    File.Delete(output_excel);

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
            string url = "https://blog.csdn.net/xcc34452366/article/details/135768052?spm=1001.2014.3001.5502";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(string model, string version, string in_data, string in_field)
        {
            List<string> result = new List<string>();

            if (in_data != "" && in_field != "")
            {
                // 检查字段值是否符合要求
                if (version == "旧版")
                {
                    if (model == "代码转名称")
                    {
                        string result_value = CheckTool.CheckFieldValue(in_data, in_field, GlobalData.dic_ydyh.Keys.ToList());
                        if (result_value != "")
                        {
                            result.Add(result_value);
                        }
                    }
                    else
                    {
                        string result_value = CheckTool.CheckFieldValue(in_data, in_field, GlobalData.dic_ydyh.Values.ToList());
                        if (result_value != "")
                        {
                            result.Add(result_value);
                        }
                    }

                }
                else
                {
                    if (model == "代码转名称")
                    {
                        string result_value = CheckTool.CheckFieldValue(in_data, in_field, GlobalData.dic_ydyh_new.Keys.ToList());
                        if (result_value != "")
                        {
                            result.Add(result_value);
                        }
                    }
                    else
                    {
                        string result_value = CheckTool.CheckFieldValue(in_data, in_field, GlobalData.dic_ydyh_new.Values.ToList());
                        if (result_value != "")
                        {
                            result.Add(result_value);
                        }
                    }
                }
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
