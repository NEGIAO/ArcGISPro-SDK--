using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing.Events;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.UI.ProWindow;
using Org.BouncyCastle.Bcpg;
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

namespace CCTool.Scripts.GHApp.KG
{
    /// <summary>
    /// Interaction logic for YDArrange.xaml
    /// </summary>
    public partial class YDArrange : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "YDArrange";

        public YDArrange()
        {
            InitializeComponent();

            try
            {
                // 初始化参数选项
                textOutFcPath.Text = BaseTool.ReadValueFromReg(toolSet, "OutFcPath");
                cb_addField.IsChecked = BaseTool.ReadValueFromReg(toolSet, "addField").ToBool();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }


        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "湘源用地整理";

        private void openOutFcButton_Click(object sender, RoutedEventArgs e)
        {
            textOutFcPath.Text = UITool.SaveDialogFeatureClass();
        }

        private void combox_fw_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fw);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147612203";
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
                string fc_path = combox_fc.ComboxText();
                string fw = combox_fw.ComboxText();

                string OutFcPath = textOutFcPath.Text;

                bool addField = (bool)cb_addField.IsChecked;

                // 判断参数是否选择完全
                if (fc_path == "" || fw == "" || OutFcPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 保存参数
                BaseTool.WriteValueToReg(toolSet, "OutFcPath", OutFcPath);
                BaseTool.WriteValueToReg(toolSet, "addField", addField);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(fc_path, OutFcPath);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10, "补充道路用地");
                    // 裁剪
                    string clip = $@"{defGDB}\clip";
                    Arcpy.Clip(fc_path, fw, clip);

                    // 擦除
                    string erase = $@"{defGDB}\erase";
                    Arcpy.Erase(fw, clip, erase);

                    // 更新
                    Arcpy.Update(clip, erase, OutFcPath);

                    // 计算字段
                    CalculateField(OutFcPath);
                    

                    
                    if (addField)
                    {
                        pw.AddMessageMiddle(5, $"字段整理");
                        // 复制模板
                        string outPath = $@"{Project.Current.HomeFolderPath}\控规空库.rar";
                        DirTool.CopyResourceRar(@"CCTool.Data.gdb.控规空库.gdb.rar", outPath);
                        // 删除重复字段
                        List<string> fields = new List<string>() { "BSM" , "YSDM", "XZQDM", "XZQMC" };
                        Arcpy.DeleteField(OutFcPath, fields);

                        // 添加字段
                        string template = $@"{Project.Current.HomeFolderPath}\控规空库.gdb\GHYD";
                        Arcpy.AddFields(OutFcPath, template);

                        // 计算字段
                        Arcpy.CalculateField(OutFcPath, "BSM", "'350524'+'0' * (10 - len(str(!OBJECTID!))-6) + str(!OBJECTID!)");     // 标识码
                        Arcpy.CalculateField(OutFcPath, "YSDM", "2090020840");     // 要素代码
                        Arcpy.CalculateField(OutFcPath, "DKBH", "!LANDINDEX!");     // 地块编号
                        Arcpy.CalculateField(OutFcPath, "YDFLMC", "!LANDNAME!");     // 用地分类名称
                        Arcpy.CalculateField(OutFcPath, "YDFLDM", "!LANDCODE!");     // 用地分类代码

                        Arcpy.CalculateField(OutFcPath, "RJLSX", "!MAXCASRAT!");     // 容积率上限
                        Arcpy.CalculateField(OutFcPath, "JZMDSX", "!MAXBUDRAT!");     // 建筑密度上限
                        Arcpy.CalculateField(OutFcPath, "LDLXX", "!MINGRNRAT!");     // 绿地率下限
                        Arcpy.CalculateField(OutFcPath, "JZGDSX", "!MAXHEIGHT!");     // 建筑高度上限
                        Arcpy.CalculateField(OutFcPath, "PTSS", "!SHARESUP!");     // 配套设施
                        Arcpy.CalculateField(OutFcPath, "BZXX", "!REMARK!");     // 备注信息

                        List<string> fields2 = new List<string>() { "UNITNAME", "LANDINDEX", "LANDCODE", "LANDNAME", "LANDAREA", "TOTALAREA", "BULDAREA", "MAXCASRAT", "MAXBUDRAT", "MINGRNRAT", "MAXHEIGHT", "SHARESUP", "CONTRLTXT", "GHQX", "REMARK", "ORIG_FID" };
                        Arcpy.DeleteField(OutFcPath, fields2);

                    }
                    // 加载
                    MapCtlTool.AddLayerToMap(OutFcPath);

                    Arcpy.Delect(clip);
                    Arcpy.Delect(erase);

                });

                pw.AddMessageEnd();

                // 删除
                try
                {
                    Arcpy.Delect($@"{Project.Current.HomeFolderPath}\控规空库.gdb");
                }
                catch (Exception)
                {

                    return;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        // 计算道路用地字段
        private void CalculateField(string fc)
        {
            Table table = fc.TargetTable();
            using RowCursor rowCursor = table.Search();
            while (rowCursor.MoveNext())
            {
                Row row = rowCursor.Current;
                string bm = row["LANDCODE"]?.ToString();
                string mc = row["LANDNAME"]?.ToString();

                if (bm == "")
                {
                    row["LANDCODE"] = "1207";
                }
                if (mc == "")
                {
                    row["LANDNAME"] = "城镇村道路用地";
                }
                row.Store();
            }
        }

        private List<string> CheckData(string fc, string outFc)
        {
            List<string> result = new List<string>();

            List<string> fields = new List<string>() { "LANDINDEX", "LANDCODE", "LANDNAME", "MAXCASRAT", "MAXBUDRAT", "MINGRNRAT", "MAXHEIGHT" };
            // 检查字段是否存在
            string result_value = CheckTool.IsHaveFieldInTarget(fc, fields);
            if (result_value != "")
            {
                result.Add(result_value);
            }

            // 检查是否是数字开头
            string result_numric = CheckTool.CheckGDBIsNumeric(outFc);
            if (result_numric != "")
            {
                result.Add(result_value);
            }
            return result;
        }

    }
}
