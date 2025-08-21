using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.GeoProcessing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using MathNet.Numerics.LinearAlgebra.Factorization;
using NPOI.OpenXmlFormats.Dml.Diagram;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using CheckBox = System.Windows.Controls.CheckBox;
using Field = ArcGIS.Core.Data.Field;
using Geometry = ArcGIS.Core.Geometry.Geometry;
using MessageBox = System.Windows.Forms.MessageBox;
using Polygon = ArcGIS.Core.Geometry.Polygon;
using Row = ArcGIS.Core.Data.Row;

namespace CCTool.Scripts.DataPross.TXT
{
    /// <summary>
    /// Interaction logic for SHP2TXTbyCom.xaml
    /// </summary>
    public partial class SHP2TXTbyCom : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "SHP2TXTbyCom";

        public SHP2TXTbyCom()
        {
            InitializeComponent();
            // 初始化
            try
            {
                // 小数位数
                combox_digit.Items.Add("1");
                combox_digit.Items.Add("2");
                combox_digit.Items.Add("3");
                combox_digit.Items.Add("4");
                combox_digit.Items.Add("5");
                combox_digit.Items.Add("6");

                combox_digit.SelectedIndex = BaseTool.ReadValueFromReg(toolSet, "digit_index").ToInt();

                // xy
                cb_xy_1.Items.Add("X");
                cb_xy_1.Items.Add("Y");
                cb_xy_1.SelectedIndex = BaseTool.ReadValueFromReg(toolSet, "xy_1").ToInt(1);

                cb_xy_2.Items.Add("X");
                cb_xy_2.Items.Add("Y");
                cb_xy_2.SelectedIndex = BaseTool.ReadValueFromReg(toolSet, "xy_2").ToInt(0);

                // 文本框信息
                inputFolder.Text = BaseTool.ReadValueFromReg(toolSet, "folder_input");
                outputFolder.Text = BaseTool.ReadValueFromReg(toolSet, "folder_output");
                txtBox_head.Text = BaseTool.ReadValueFromReg(toolSet, "txt_head");

                string txtJ = BaseTool.ReadValueFromReg(toolSet, "txt_J");
                txt_J.Text = txtJ == ""?"J":txtJ;

                string txtEnd = BaseTool.ReadValueFromReg(toolSet, "txt_end");
                txt_end.Text = txtEnd == "" ? "@" : txtEnd;

                // 其它参数
                check_startPoint.IsChecked = BaseTool.ReadValueFromReg(toolSet, "start_point").ToBool("false");
                check_closed.IsChecked = BaseTool.ReadValueFromReg(toolSet, "closed").ToBool("true");
                check_Part_Reshot.IsChecked = BaseTool.ReadValueFromReg(toolSet, "part_reshot").ToBool("false");

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "SHP转TXT_通用版";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 文件路径
                string folder_input = inputFolder.Text;
                string folder_output = outputFolder.Text;
                // 要处理的shp
                var cb_shps = listbox_shp.Items;

                // txt参数
                string txtHead = txtBox_head.Text;
                string txtJ = txt_J.Text;
                string txtEnd = txt_end.Text;

                int digit_xy = int.Parse(combox_digit.Text);
                string xy_1 = cb_xy_1.Text;
                string xy_2 = cb_xy_2.Text;

                bool startPoint = (bool)check_startPoint.IsChecked;
                bool isClosed = (bool)check_closed.IsChecked;
                bool isRePart = (bool)check_Part_Reshot.IsChecked;

                // 可选字段
                List<string> cbs = new List<string>()
                {
                    cb_1.Text, cb_2.Text,cb_3.Text, cb_4.Text, cb_5.Text,cb_6.Text, cb_7.Text, cb_8.Text
                };

                // 判断参数是否选择完全
                if (folder_input == "" || folder_output == "" || cb_shps.Count == 0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 参数写入本地
                BaseTool.WriteValueToReg(toolSet, "folder_input", folder_input);
                BaseTool.WriteValueToReg(toolSet, "folder_output", folder_output);

                BaseTool.WriteValueToReg(toolSet, "txt_head", txtHead);
                BaseTool.WriteValueToReg(toolSet, "txt_J", txtJ);
                BaseTool.WriteValueToReg(toolSet, "txt_end", txtEnd);

                BaseTool.WriteValueToReg(toolSet, "digit_index", combox_digit.SelectedIndex);
                BaseTool.WriteValueToReg(toolSet, "xy_1", cb_xy_1.SelectedIndex);
                BaseTool.WriteValueToReg(toolSet, "xy_2", cb_xy_2.SelectedIndex);

                BaseTool.WriteValueToReg(toolSet, "start_point", startPoint);
                BaseTool.WriteValueToReg(toolSet, "closed", isClosed);
                BaseTool.WriteValueToReg(toolSet, "part_reshot", isRePart);


                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 如果是空路径，就创建输出路径
                if (!Directory.Exists(folder_output))
                {
                    Directory.CreateDirectory(folder_output);
                }

                // 获取所有选中的shp
                List<string> list_shpPath = new List<string>();
                foreach (CheckBox shp in cb_shps)
                {
                    if (shp.IsChecked == true)
                    {
                        list_shpPath.Add(folder_input + shp.Content);
                    }
                }

                pw.AddMessageStart("获取参数");

                await QueuedTask.Run(async () =>
                {
                    foreach (string fullPath in list_shpPath)
                    {
                        // 初始化写入txt的内容
                        string txt_all = txtHead + "\r\n" + "[地块坐标]" + "\r\n";

                        pw.AddMessageMiddle(10, fullPath);
                        string shp_name = fullPath[(fullPath.LastIndexOf(@"\") + 1)..];  // 获取要素名
                        string shp_path = fullPath[..(fullPath.LastIndexOf(@"\"))];  // 获取shp名

                        // 打开shp
                        FileSystemConnectionPath fileConnection = new FileSystemConnectionPath(new Uri(shp_path), FileSystemDatastoreType.Shapefile);
                        using FileSystemDatastore shapefile = new FileSystemDatastore(fileConnection);
                        // 获取FeatureClass
                        FeatureClass featureClass = shapefile.OpenDataset<FeatureClass>(shp_name);

                        using (RowCursor rowCursor = featureClass.Search())
                        {
                            while (rowCursor.MoveNext())
                            {
                                using Feature feature = rowCursor.Current as Feature;
                                // 拼合抬头行
                                string title = await MergeHeadString(feature, txtEnd,cbs);
                                txt_all += title;

                                // 获取面要素的JSON文字
                                Polygon polygon = feature.GetShape() as Polygon;
                                // 是否重排，从西北角开始
                                List<List<MapPoint>> mapPoints = new List<List<MapPoint>>();
                                if (startPoint)
                                {
                                    mapPoints = polygon.ReshotMapPoint(true);
                                }
                                else
                                {
                                    mapPoints = polygon.MapPointsFromPolygon(true);
                                }

                                // 加一行title
                                int count = 0;    // 点的个数
                                foreach (var points in mapPoints)
                                {
                                    count += points.Count;
                                }


                                int index = 1;   // 点号
                                int lastNum = 1;
                                for (int i = 0; i < mapPoints.Count; i++)
                                {
                                    for (int j = 0; j < mapPoints[i].Count; j++)
                                    {
                                        // 写入点号
                                        // 小数位数补齐0
                                        string XX = mapPoints[i][j].X.RoundWithFill(digit_xy);
                                        string YY = mapPoints[i][j].Y.RoundWithFill(digit_xy);

                                        // 点号计算
                                        int ptIndex = index;
                                        // 如果是最后一个点
                                        if (j == mapPoints[i].Count - 1)
                                        {
                                            if (isRePart)
                                            {
                                                ptIndex = 1;
                                            }
                                            else
                                            {
                                                ptIndex = lastNum;
                                            }
                                        }

                                        // 写入折点的XY值，如果首末点不重合，最后一点就不写入
                                        if (!(isClosed == false && j == mapPoints[i].Count - 1))
                                        {
                                            // 获取XY值
                                            string xyValue_1 = xy_1=="X"?XX:YY;
                                            string xyValue_2 = xy_2 == "X" ? XX : YY;

                                            // 加入文本
                                            txt_all += $"{txtJ}{ptIndex},{i + 1},{xyValue_1},{xyValue_2}\r\n";
                                        }

                                        // 如果不是最后一个点，增加点号
                                        if (j != mapPoints[i].Count - 1)
                                        {
                                            index++;

                                        }
                                        else   // 如果是最后一个点
                                        {
                                            lastNum += (mapPoints[i].Count - 1);

                                            if (isRePart)    // 新部件点号重新从1开始
                                            {
                                                index = 1;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        // 写入txt文件
                        string txtPath = @$"{folder_output}\{shp_name.Replace(".shp", "")}.txt";
                        if (File.Exists(txtPath))
                        {
                            File.Delete(txtPath);
                        }
                        File.WriteAllText(txtPath, txt_all);
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

        // 拼合抬头行
        private async Task<string> MergeHeadString(Feature feature, string txtEnd, List<string> cbs)
        {
            string result;
            // 字段值获取
            string field1 = await GetFieldValue(feature, cbs[0]);
            string field2 = await GetFieldValue(feature, cbs[1]);
            string field3 = await GetFieldValue(feature, cbs[2]);
            string field4 = await GetFieldValue(feature, cbs[3]);
            string field5 = await GetFieldValue(feature, cbs[4]);
            string field6 = await GetFieldValue(feature, cbs[5]);
            string field7 = await GetFieldValue(feature, cbs[6]);
            string field8 = await GetFieldValue(feature, cbs[7]);

            // 拼合
            result = $"{field1},{field2},{field3},{field4},{field5},{field6},{field7},{field8},{txtEnd}" + "\r\n";
            return result;
        }

        // 获取字段值
        private async Task<string> GetFieldValue(Feature feature, string fieldString)
        {
            string result;

            List<string> list = await GetFieldsName();

            // 如果是取自字段
            if (list.Contains(fieldString))
            {
                result = feature[fieldString]?.ToString();
            }
            // 如果是手输的文本
            else
            {
                result = fieldString;
            }

            return result;
        }


        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            // 打开SHP文件夹
            string folder = UITool.OpenDialogFolder();
            inputFolder.Text = folder;
            // 更新shp列表框
            UpdataListboxSHP(folder);
        }

        // 更新shp列表框
        private void UpdataListboxSHP(string folder)
        {
            // 清除listbox
            listbox_shp.Items.Clear();
            // 生成SHP要素列表
            if (folder != "" && Directory.Exists(folder))
            {
                // 获取所有shp文件
                var files = DirTool.GetAllFiles(folder, ".shp");
                foreach (var file in files)
                {
                    // 将shp文件做成checkbox放入列表中
                    CheckBox cb = new CheckBox();
                    cb.Content = file.Replace(folder, "");
                    cb.IsChecked = true;
                    listbox_shp.Items.Add(cb);
                }
            }
        }


        private void openTXTFolderButton_Click(object sender, RoutedEventArgs e)
        {
            outputFolder.Text = UITool.OpenDialogFolder();
        }


        // 获取所有要素共有的文本字段
        public async Task<List<string>> GetFieldsName(string FieldType = "all")
        {
            List<string> fieldsName = new List<string>();

            // 获取所有选中的shp
            List<string> shpList = new List<string>();

            // 在UI线程上执行添加item的操作
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                string folder_path = inputFolder.Text;
                var cb_shps = listbox_shp.Items;

                foreach (CheckBox shp in cb_shps)
                {
                    if (shp.IsChecked == true)
                    {
                        shpList.Add(folder_path + shp.Content);
                    }
                }
            });

            // 定义一个dic
            Dictionary<string, int> numbers = new Dictionary<string, int>();

            foreach (var shp in shpList)
            {
                // 获取所选图层的所有字段
                var fields = await QueuedTask.Run(() =>
                {
                    return GisTool.GetFieldsNameFromTarget(shp, FieldType);
                });

                // 查看字段名
                foreach (string field in fields)
                {
                    if (numbers.ContainsKey(field))
                    {
                        numbers[field] += 1;
                    }
                    else
                    {
                        numbers.Add(field, 1);
                    }
                }
            }

            // 检查一下，个数不足的去除
            int shpCount = shpList.Count;   // shp的总个数
            foreach (var number in numbers)
            {
                if (number.Value == shpCount)
                {
                    fieldsName.Add(number.Key);
                }
            }

            return fieldsName;
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/145516105";
            UITool.Link2Web(url);
        }



        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_shp);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_shp);
        }

        private async void cb_1_Open(object sender, EventArgs e)
        {
            List<string> list = await GetFieldsName();
            UITool.AddStringListToCombox(list, cb_1);
        }

        private async void cb_2_Open(object sender, EventArgs e)
        {
            List<string> list = await GetFieldsName();
            UITool.AddStringListToCombox(list, cb_2);
        }

        private async void cb_3_Open(object sender, EventArgs e)
        {
            List<string> list = await GetFieldsName();
            UITool.AddStringListToCombox(list, cb_3);
        }

        private async void cb_4_Open(object sender, EventArgs e)
        {
            List<string> list = await GetFieldsName();
            UITool.AddStringListToCombox(list, cb_4);
        }

        private async void cb_5_Open(object sender, EventArgs e)
        {
            List<string> list = await GetFieldsName();
            UITool.AddStringListToCombox(list, cb_5);
        }

        private async void cb_6_Open(object sender, EventArgs e)
        {
            List<string> list = await GetFieldsName();
            UITool.AddStringListToCombox(list, cb_6);
        }

        private async void cb_7_Open(object sender, EventArgs e)
        {
            List<string> list = await GetFieldsName();
            UITool.AddStringListToCombox(list, cb_7);
        }

        private async void cb_8_Open(object sender, EventArgs e)
        {
            List<string> list = await GetFieldsName();
            UITool.AddStringListToCombox(list, cb_8);
        }

        private void listbox_shp_Load(object sender, RoutedEventArgs e)
        {
            // 打开SHP文件夹
            string folder = inputFolder.Text;
            // 更新shp列表框
            UpdataListboxSHP(folder);
        }
    }
}
