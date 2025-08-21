using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.SS.Formula.PTG;
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

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    /// <summary>
    /// Interaction logic for Json2Polygon.xaml
    /// </summary>
    public partial class Json2Polygon : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "Json2Polygon";

        public Json2Polygon()
        {
            InitializeComponent();

            // combox_sr框中添加几种预制坐标系
            combox_sr.Items.Add("GCS_WGS_1984");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_25");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_26");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_27");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_28");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_29");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_30");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_31");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_32");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_33");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_34");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_35");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_36");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_37");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_38");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_39");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_40");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_41");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_42");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_43");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_44");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_Zone_45");

            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_75E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_78E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_81E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_84E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_87E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_90E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_93E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_96E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_99E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_102E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_105E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_108E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_111E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_114E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_117E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_120E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_123E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_126E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_129E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_132E");
            combox_sr.Items.Add("CGCS2000_3_Degree_GK_CM_135E");

            combox_sr.SelectedIndex = 0;

            // 初始化其它参数选项
            txtFCPath.Text = BaseTool.ReadValueFromReg(toolSet, "fc_path");
        }


        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "类Json文本转面要素";



        private void combox_field_DropDown(object sender, EventArgs e)
        {
            string fc = combox_fc.ComboxText();
            UITool.AddTextFieldsToComboxPlus(fc, combox_field);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayerAndTableToComboxPlus(combox_fc);
        }

        private async void combox_fc_Closed(object sender, EventArgs e)
        {

            // 把所有字段放入列表框
            string fc = combox_fc.ComboxText();

            var fields = await QueuedTask.Run(() =>
            {
                return GisTool.GetFieldsNameFromTarget(fc);
            });

            foreach (var field in fields)
            {
                // 将txt文件做成checkbox放入列表中
                CheckBox cb = new CheckBox();
                cb.Content = field;
                cb.IsChecked = true;
                listbox_field.Items.Add(cb);
            }
        }


        private void openFCButton_Click(object sender, RoutedEventArgs e)
        {
            txtFCPath.Text = UITool.SaveDialogFeatureClass();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string fc_path = txtFCPath.Text;
                string in_table = combox_fc.ComboxText();
                string field = combox_field.ComboxText();

                string spatial_reference = combox_sr.Text;

                var cb_txts = listbox_field.Items;
                // 获取选定的字段列表
                List<string> addFields = new List<string>();
                foreach (CheckBox cb_txt in cb_txts)
                {
                    if (cb_txt.IsChecked == true)
                    {
                        addFields.Add(cb_txt.Content.ToString());
                    }
                }

                // 判断参数是否选择完全
                if (fc_path == "" || in_table == "" || field == "" || spatial_reference == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "fc_path", fc_path);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                // 异步执行
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(fc_path);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    string gdbPath = fc_path[..(fc_path.LastIndexOf(".gdb") + 4)];
                    string fcName = fc_path[(fc_path.LastIndexOf(@"\") + 1)..];


                    pw.AddMessageMiddle(0,"创建空要素类");
                    // 创建空要素类
                    Arcpy.CreateFeatureclass(gdbPath, fcName, "POLYGON", spatial_reference);

                    pw.AddMessageMiddle(10, "添加字段");
                    // 添加字段
                    if (cb_txts.Count > 0)
                    {
                        foreach (string addField in addFields)
                        {
                            GisTool.AddField(fc_path, addField, FieldType.String);
                        }
                    }

                    pw.AddMessageMiddle(10, "解析文本并创建要素");

                    // 打开数据库
                    using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbPath))))
                    {
                        // 创建要素并添加到要素类中
                        using FeatureClass featureClass = gdb.OpenDataset<FeatureClass>(fcName);

                        // 解析文本并创建要素
                        using Table table = in_table.TargetTable();
                        using RowCursor rowCursor = table.Search();
                        while (rowCursor.MoveNext())
                        {
                            var mapPoints = new List<Coordinate2D>(); ;

                            using Row row = rowCursor.Current;
                            var re = row[field];
                            if (re is not null)
                            {
                                // 获取目标文本，并清空空格
                                string text = re.ToString().Replace(" ", "");
                                // 提取点集文本
                                string target = text.Split("],Polygon]")[0][2..];
                                // 考虑是否有多部件
                                string[] parts = target.Split("]],");

                                foreach (string part in parts)
                                {
                                    string text2 = part;
                                    // 如果有多部件，就补充一下被分割的文本
                                    if (parts.Length > 1 && !text2.EndsWith("]]"))
                                    {
                                        text2 += "]]";
                                    }

                                    // 再分割成点集
                                    string[] pts = text2[1..^1].Split("],");

                                    // 添加到点集中
                                    for (int i = 0; i < pts.Length; i++)
                                    {
                                        string tt = pts[i];
                                        if (i != pts.Length - 1)    // 最后一个
                                        {
                                            tt += "]";
                                        }

                                        double xx = double.Parse(tt[1..^1].Split(',')[0]);
                                        double yy = double.Parse(tt[1..^1].Split(',')[1]);
                                        mapPoints.Add(new Coordinate2D(xx, yy));
                                    }
                                }
                            }

                            /// 构建面要素
                            // 创建编辑操作对象
                            EditOperation editOperation = new EditOperation();
                            editOperation.Callback(context =>
                            {
                                // 获取要素定义
                                FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();
                                // 创建RowBuffer
                                using RowBuffer rowBuffer = featureClass.CreateRowBuffer();

                                // 写入字段值
                                foreach (string addField in addFields)
                                {
                                    rowBuffer[addField] = row[addField]?.ToString();
                                }

                                PolygonBuilderEx pb = new PolygonBuilderEx(mapPoints);

                                // 给新添加的行设置形状
                                rowBuffer[featureClassDefinition.GetShapeField()] = pb.ToGeometry();

                                // 在表中创建新行
                                using Feature feature = featureClass.CreateRow(rowBuffer);
                                context.Invalidate(feature);      // 标记行为无效状态
                            }, featureClass);

                            // 执行编辑操作
                            editOperation.Execute();
                        }
                    }

                    // 保存编辑
                    Project.Current.SaveEditsAsync();

                    // 加载图层
                    MapCtlTool.AddLayerToMap(fc_path);

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
            string url = "https://blog.csdn.net/xcc34452366/article/details/145092497";
            UITool.Link2Web(url);
        }

        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_field);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_field);
        }

        private List<string> CheckData(string fcPath)
        {
            List<string> result = new List<string>();

            string result_value = CheckTool.CheckGDBFeature(fcPath);
            if (result_value != "")
            {
                result.Add(result_value);
            }

            return result;
        }
    }
}
