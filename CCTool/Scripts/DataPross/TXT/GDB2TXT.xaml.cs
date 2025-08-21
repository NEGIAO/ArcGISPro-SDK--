using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
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
    /// Interaction logic for GDB2TXT.xaml
    /// </summary>
    public partial class GDB2TXT : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "GDB2TXT";

        public GDB2TXT()
        {
            InitializeComponent();

            try
            {
                // 定义一个更新title的事件
                EventCenter.AddListener(EventDefine.UpdataTitle, UpdataTitle);

                combox_digit.Items.Add("1");
                combox_digit.Items.Add("2");
                combox_digit.Items.Add("3");
                combox_digit.Items.Add("4");
                combox_digit.Items.Add("5");
                combox_digit.Items.Add("6");
                combox_digit.SelectedIndex = 2;

                // 初始化文本框信息
                UpdataTitle();

                txtFolder.Text = BaseTool.ReadValueFromReg(toolSet, "folder_path");
                txtFolder2.Text = BaseTool.ReadValueFromReg(toolSet, "folder_txt");
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message+ ee.StackTrace);
                return;
            }
            
            
        }

        private void UpdataTitle()
        {
            string initTitle = BaseTool.ReadValueFromReg("TitleBox", "initTitle");
            txtBox_head.Text = initTitle;
        }

        // 定义Title框
        private TitleMessage titleMessage = null;

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "SHP转TXT";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string folder_path = txtFolder.Text;
                string folder_txt = txtFolder2.Text;
                var cb_shps = listbox_shp.Items;
                string field_mc = combox_mc.ComboxText();
                string field_yt = combox_yt.ComboxText();
                string field_tf = combox_tf.ComboxText();

                int digit_xy = int.Parse(combox_digit.Text);
                string txt_head = txtBox_head.Text;

                bool xyReserve = (bool)check_xy.IsChecked;
                bool haveJ = (bool)check_xy_J.IsChecked;
                bool startPoint = (bool)check_startPoint.IsChecked;
                bool isClosed = (bool)check_closed.IsChecked;
                bool isRePart = (bool)check_Part_Reshot.IsChecked;


                string field_mj = combox_mj.ComboxText();
                string field_bh = combox_time.ComboxText();

                // 判断参数是否选择完全
                if (folder_path == "" || folder_txt == "" || cb_shps.Count == 0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                BaseTool.WriteValueToReg(toolSet, "folder_path", folder_path);
                BaseTool.WriteValueToReg(toolSet, "folder_txt", folder_txt);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 如果是空路径，就创建输出路径
                if (!Directory.Exists(folder_txt))
                {
                    Directory.CreateDirectory(folder_txt);
                }

                // 获取所有选中的shp
                List<string> list_shpPath = new List<string>();
                foreach (CheckBox shp in cb_shps)
                {
                    if (shp.IsChecked == true)
                    {
                        list_shpPath.Add(folder_path + shp.Content);
                    }
                }

                pw.AddMessageStart("获取参数");

                await QueuedTask.Run(() =>
                {
                    foreach (string fullPath in list_shpPath)
                    {
                        // 初始化写入txt的内容
                        string txt_all = txt_head + "\r\n" + "[地块坐标]" + "\r\n";

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
                                using (Feature feature = rowCursor.Current as Feature)
                                {
                                    // 获取地块名称，地块性质
                                    Row row = feature as Row;

                                    string ft_mc = "";
                                    if (field_mc != "") { ft_mc = row[field_mc]?.ToString() ?? ""; }
                                    string ft_yt = "";
                                    if (field_yt != "") { ft_yt = row[field_yt]?.ToString() ?? ""; }
                                    string ft_tf = "";
                                    if (field_tf != "") { ft_tf = row[field_tf]?.ToString() ?? ""; }
                                    string ft_mj = "";
                                    if (field_mj != "") { ft_mj = double.Parse(row[field_mj].ToString()).RoundWithFill(4) ?? ""; }
                                    string ft_bh = "";
                                    if (field_bh != "") { ft_bh = row[field_bh]?.ToString() ?? ""; }

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
                                    string title = $"{count},{ft_mj},{ft_bh},{ft_mc},面,{ft_tf},{ft_yt},地上,@" + "\r\n";
                                    txt_all += title;


                                    int index = 1;   // 点号
                                    int lastNum = 1;
                                    for (int i = 0; i < mapPoints.Count; i++)
                                    {
                                        for (int j = 0; j < mapPoints[i].Count; j++)
                                        {
                                            // 写入点号
                                            // J前缀
                                            string jFront = haveJ switch
                                            {
                                                true => "J",
                                                false => "",
                                            };
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
                                            if (!(isClosed==false && j == mapPoints[i].Count - 1))
                                            {
                                                if (xyReserve)
                                                {
                                                    // 加入文本
                                                    txt_all += $"{jFront}{ptIndex},{i + 1},{XX},{YY}\r\n";
                                                }
                                                else
                                                {
                                                    // 加入文本
                                                    txt_all += $"{jFront}{ptIndex},{i + 1},{YY},{XX}\r\n";
                                                }
                                            }

                                            // 如果不是最后一个点，增加点号
                                            if (j != mapPoints[i].Count - 1)
                                            {
                                                index++;

                                            }
                                            else   // 如果是最后一个点
                                            {
                                                lastNum += (mapPoints[i].Count-1);

                                                if (isRePart)    // 新部件点号重新从1开始
                                                {
                                                    index = 1;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        // 写入txt文件
                        string txtPath = @$"{folder_txt}\{shp_name.Replace(".shp", "")}.txt";
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

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            // 打开SHP文件夹
            string folder = UITool.OpenDialogFolder();
            txtFolder.Text = folder;

            // 清除listbox
            listbox_shp.Items.Clear();
            // 生成SHP要素列表
            if (txtFolder.Text != "")
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

        private async void combox_mc_Open(object sender, EventArgs e)
        {
            try
            {
                // 获取共有字段
                List<string> list = await GetFieldsName();
                // 加到combox中
                UITool.AddStringToComboxPlus(list, combox_mc);

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        private async void combox_yt_Open(object sender, EventArgs e)
        {
            // 获取共有字段
            List<string> list = await GetFieldsName();
            // 加到combox中
            UITool.AddStringToComboxPlus(list, combox_yt);
        }


        private void openTXTFolderButton_Click(object sender, RoutedEventArgs e)
        {
            txtFolder2.Text = UITool.OpenDialogFolder();
        }

        private async void combox_mj_Open(object sender, EventArgs e)
        {
            // 获取共有字段
            List<string> list = await GetFieldsName("float_all");
            // 加到combox中
            UITool.AddFloatToComboxPlus(list, combox_mj);
        }

        private async void combox_time_Open(object sender, EventArgs e)
        {
            // 获取共有字段
            List<string> list = await GetFieldsName();
            // 加到combox中
            UITool.AddStringToComboxPlus(list, combox_time);
        }

        // 获取所有要素共有的文本字段
        public async Task<List<string>> GetFieldsName(string FieldType = "text")
        {
            List<string> fieldsName = new List<string>();

            string folder_path = txtFolder.Text;
            var cb_shps = listbox_shp.Items;

            // 获取所有选中的shp
            List<string> shpList = new List<string>();
            foreach (CheckBox shp in cb_shps)
            {
                if (shp.IsChecked == true)
                {
                    shpList.Add(folder_path + shp.Content);
                }
            }

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




        private void btn_read_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 打开进度框
                TitleMessage tm = UITool.OpenTitleWindow(titleMessage);
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        // 关闭窗口
        private void fm_Unloaded(object sender, RoutedEventArgs e)
        {
            // 更新当前文本框至配置文件
            string txt = txtBox_head.Text;

            BaseTool.WriteValueToReg("TitleBox", "initTitle", txt);

            // 关闭窗口时移除监听事件
            EventCenter.RemoveListener(EventDefine.UpdataTitle, UpdataTitle);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/140746501?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        private async void combox_tf_Open(object sender, EventArgs e)
        {
            // 获取共有字段
            List<string> list = await GetFieldsName();
            // 加到combox中
            UITool.AddStringToComboxPlus(list, combox_tf);
        }
    }
}
