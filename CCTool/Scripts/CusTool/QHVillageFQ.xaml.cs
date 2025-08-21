using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
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

namespace CCTool.Scripts.CusTool
{
    /// <summary>
    /// Interaction logic for QHVillageFQ.xaml
    /// </summary>
    public partial class QHVillageFQ : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public QHVillageFQ()
        {
            InitializeComponent();
            // 初始化
            Init();
        }
        // 初始化
        public void Init()
        {
            try
            {
                // 将剩余要素图层放在目标图层中
                UITool.AddFeatureLayersAndTablesToListbox(listbox_targetFeature);

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }


        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "青海村规分区";

        private void combox_fc_gh_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void combox_bmField_DropDown(object sender, EventArgs e)
        {
            // 将图层字段加入到Combox列表中
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_bmField);
        }

        private void openFeatureClassButton_Click(object sender, RoutedEventArgs e)
        {
            // 打开Excel文件
            string path = UITool.SaveDialogFeatureClass();
            // 将Excel文件的路径
            textFeatureClassPath.Text = path;
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "";
            UITool.Link2Web(url);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                string init_gdb = Project.Current.DefaultGeodatabasePath;
                string init_foder = Project.Current.HomeFolderPath;

                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string field_bm = combox_bmField.ComboxText();
                // 输入最小面积
                string miniAreaText = textMiniArea.Text == "" ? "0" : textMiniArea.Text;
                double miniArea;
                try
                {
                    miniArea = double.Parse(miniAreaText);
                }
                catch (Exception)
                {
                    MessageBox.Show("请输入一个正确的面积值！");
                    return;
                }

                string output_fq = textFeatureClassPath.Text;
                // 其它分区图层
                List<string> targetFeatureClasses = UITool.GetCheckboxStringFromListBox(listbox_targetFeature);
                // 去除路径结构
                List<string> targetFCNames = new List<string>();
                foreach (string targetClass in targetFeatureClasses)
                {
                    int index = targetClass.IndexOf(@"\");
                    targetFCNames.Add(targetClass[(index + 1)..]);
                }

                // 判断参数是否选择完全
                if (fc_path == "" || field_bm == "" || output_fq == "")
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
                    // 复制映射表
                    string excelMapper = "【青海】用地用海代码_村庄三区";
                    string mapper = @$"{init_foder}\{excelMapper}.xlsx";
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.{excelMapper}.xlsx", mapper);

                    pw.AddMessageStart("农业、生态、建设三区归类");
                    // 复制要素
                    string fcCopy = @$"{init_gdb}\青海分区";
                    Arcpy.CopyFeatures(fc_path, fcCopy);
                    // 添加字段
                    string gnField = "村庄分区";
                    Arcpy.AddField(fcCopy, gnField, "TEXT");
                    // 村庄功能映射
                    ComboTool.AttributeMapper(fcCopy, field_bm, gnField, $@"{mapper}\sheet1$");

                    // 其它分区叠加
                    string updataFC = @$"{init_gdb}\青海分区更新";
                    if (targetFeatureClasses.Count > 0)
                    {
                        pw.AddMessageMiddle(20, "其它分区叠加");
                        for (int i = 0; i < targetFeatureClasses.Count; i++)
                        {
                            pw.AddMessageMiddle(10, targetFCNames[i], Brushes.Gray);
                            string temFC = $@"{init_gdb}\{targetFCNames[i]}_tem";
                            // 复制要素
                            Arcpy.CopyFeatures(targetFeatureClasses[i], temFC);
                            // 添加字段
                            Arcpy.AddField(temFC, gnField, "TEXT");
                            // 计算字段字段
                            Arcpy.CalculateField(temFC, gnField, $"'{targetFCNames[i]}'");
                            // 更新
                            Arcpy.Update(fcCopy, temFC, updataFC);
                            Arcpy.CopyFeatures(updataFC, fcCopy);
                            // 删除中间数据
                            Arcpy.Delect(temFC);
                        }

                    }

                    pw.AddMessageMiddle(10, $"融合");
                    // 融合
                    string fcCopy2 = @$"{init_gdb}\青海分区2";
                    Arcpy.Dissolve(fcCopy, fcCopy2, gnField);
                    string fcCopy3 = @$"{init_gdb}\青海分区3";
                    Arcpy.MultipartToSinglepart(fcCopy2, fcCopy3);

                    pw.AddMessageMiddle(10, $"融合小图斑({miniArea}平方米以下)");
                    // 融合小图斑
                    if (miniArea > 0)
                    {
                        ComboTool.FeatureClassEliminate(fcCopy3, output_fq, $"SHAPE_Area < {miniArea}");
                    }

                    // 加载
                    MapCtlTool.AddLayerToMap(output_fq);

                    // 删除中间数据
                    pw.AddMessageMiddle(10, "删除中间数据");
                    Arcpy.Delect(fcCopy);
                    Arcpy.Delect(fcCopy2);
                    Arcpy.Delect(fcCopy3);
                    Arcpy.Delect(updataFC);

                    File.Delete(mapper);
                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void combox_fc_DropClosed(object sender, EventArgs e)
        {
            try
            {
                // 获取默认数据库
                string init_gdb = Project.Current.DefaultGeodatabasePath;
                string lyName = combox_fc.ComboxText();
                string lyNameSingle = lyName[(lyName.LastIndexOf(@"\") + 1)..];
                textFeatureClassPath.Text = $@"{init_gdb}\{lyNameSingle}_转分区";

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }
    }
}
