using ActiproSoftware.Windows.Shapes;
using ApeFree.DataStore;
using ApeFree.DataStore.Core;
using ApeFree.DataStore.Local;
using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Windows;
using CCTool.Scripts.UI.ProWindow;
using NPOI.POIFS.Crypt.Dsig;
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

namespace CCTool.Scripts.DataPross.TXT
{
    /// <summary>
    /// Interaction logic for FeatureClass2TXT.xaml
    /// </summary>
    public partial class FeatureClass2TXT : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "FeatureClass2TXT";
        public FeatureClass2TXT()
        {
            InitializeComponent();

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

            txtFolder2.Text = BaseTool.ReadValueFromReg(toolSet, "folder_txt");
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
        string tool_name = "要素类转TXT";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string lyName = combox_fc.ComboxText();
                string folder_txt = txtFolder2.Text;
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

                bool isMerge = (bool)checkbox_merge.IsChecked;

                string field_mj = combox_mj.ComboxText();
                string field_bh = combox_time.ComboxText();

                // 判断参数是否选择完全
                if (lyName == "" || folder_txt == "" || field_mc == "" || field_bh == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

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

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(lyName, field_mc);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    // 获取目标FeatureLayer
                    FeatureLayer featurelayer = lyName.TargetFeatureLayer();

                    // 遍历面要素类中的所有要素(有选择就按选择)
                    RowCursor rowCursor = featurelayer.TargetSelectCursor();

                    // txt起始段
                    string txt_title = txt_head + "\r\n" + "[地块坐标]" + "\r\n";
                    // 要素的txt
                    List<string> txt_feature_list = new List<string>();
                    while (rowCursor.MoveNext())
                    {
                        using Feature feature = rowCursor.Current as Feature;
                        // 获取单个要素的txt
                        string txt_feature = GetWordFromFeature(feature, field_mc, field_yt, field_tf, field_mj, field_bh, haveJ, digit_xy, xyReserve, startPoint, isClosed, isRePart);
                        txt_feature_list.Add(txt_feature);

                        // 如果分多个txt
                        if (!isMerge)
                        {
                            // 单个要素的txt_all
                            string txt_all = txt_title + txt_feature;
                            // 写入txt文件
                            string txtPath = @$"{folder_txt}\{feature[field_bh]}.txt";

                            if (File.Exists(txtPath))
                            {
                                File.Delete(txtPath);
                            }
                            File.WriteAllText(txtPath, txt_all);
                        }
                    }

                    // 如果合并成一个txt
                    if (isMerge)
                    {
                        string txt_all = txt_title;
                        foreach (var txt in txt_feature_list)
                        {
                            txt_all += txt;
                        }
                        // 写入txt文件
                        string txtPath = @$"{folder_txt}\{lyName}.txt";

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

        // 从要素中获取属性文字
        public string GetWordFromFeature(Feature feature, string field_mc, string field_yt, string field_tf, string field_mj, string field_bh, bool haveJ, int digit_xy, bool xyReserve, bool startPoint, bool isClosed, bool isRePart)
        {
            string txt_feature = "";

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
            ArcGIS.Core.Geometry.Polygon polygon = feature.GetShape() as ArcGIS.Core.Geometry.Polygon;
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
            txt_feature += title;

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
                    if (!(isClosed == false && j == mapPoints[i].Count - 1))
                    {
                        if (xyReserve)
                        {
                            // 加入文本
                            txt_feature += $"{jFront}{ptIndex},{i + 1},{XX},{YY}\r\n";
                        }
                        else
                        {
                            // 加入文本
                            txt_feature += $"{jFront}{ptIndex},{i + 1},{YY},{XX}\r\n";
                        }
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
            return txt_feature;
        }


        private void combox_mc_Open(object sender, EventArgs e)
        {
            string lyName = combox_fc.ComboxText();
            UITool.AddTextFieldsToComboxPlus(lyName, combox_mc);

        }

        private void combox_yt_Open(object sender, EventArgs e)
        {
            string lyName = combox_fc.ComboxText();
            UITool.AddTextFieldsToComboxPlus(lyName, combox_yt);
        }


        private void openTXTFolderButton_Click(object sender, RoutedEventArgs e)
        {
            txtFolder2.Text = UITool.OpenDialogFolder();
        }

        private void combox_mj_Open(object sender, EventArgs e)
        {
            string lyName = combox_fc.ComboxText();
            UITool.AddFloatFieldsToComboxPlus(lyName, combox_mj);
        }

        private void combox_time_Open(object sender, EventArgs e)
        {
            string lyName = combox_fc.ComboxText();
            UITool.AddTextFieldsToComboxPlus(lyName, combox_time);
        }

        private void combox_fc_DropOpen(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
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

        private List<string> CheckData(string in_data , string in_field)
        {
            List<string> result = new List<string>();

            if (in_data != "" && in_field != "")
            {

                // 检查字段值是否为空
                string fieldEmptyResult = CheckTool.CheckFieldValueSpace(in_data, in_field);
                if (fieldEmptyResult != "")
                {
                    result.Add(fieldEmptyResult);
                }
            }

            return result;
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/140745914?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        private void combox_tf_Open(object sender, EventArgs e)
        {
            string lyName = combox_fc.ComboxText();
            UITool.AddTextFieldsToComboxPlus(lyName, combox_tf);
        }
    }
}
