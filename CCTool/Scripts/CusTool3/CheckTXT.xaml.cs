using ActiproSoftware.Windows.Extensions;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Data.Exceptions;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells.Drawing;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Windows;
using NPOI.OpenXmlFormats.Shared;
using NPOI.OpenXmlFormats.Vml;
using NPOI.POIFS.Crypt;
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
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;
using static System.Net.Mime.MediaTypeNames;
using CheckBox = System.Windows.Controls.CheckBox;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for CheckTXT.xaml
    /// </summary>
    public partial class CheckTXT : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "CheckTXT";
        public CheckTXT()
        {
            InitializeComponent();

            // 初始化
            textFolderPath.Text = BaseTool.ReadValueFromReg(toolSet, "folderPath");
            textWordPath.Text = BaseTool.ReadValueFromReg(toolSet, "wordPath");

            // 更新列表框
            UpdataListBox();
        }



        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "TXT质检(吉)";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string folderPath = textFolderPath.Text;
                string wordPath = textWordPath.Text;
                bool isReverse = (bool)cb_reverse.IsChecked;

                var cb_txts = listbox_txt.Items;

                // 参数写入本地
                BaseTool.WriteValueToReg(toolSet, "folderPath", folderPath);
                BaseTool.WriteValueToReg(toolSet, "wordPath", wordPath);

                // 判断参数是否选择完全
                if (folderPath == "" || wordPath == "" || cb_txts.Count == 0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 获取所有选中的txt
                List<string> list_txtPath = new List<string>();
                foreach (CheckBox shp in cb_txts)
                {
                    if (shp.IsChecked == true)
                    {
                        list_txtPath.Add(folderPath + shp.Content);
                    }
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                // 异步执行
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");

                    // 错误结果
                    string errResult = "";

                    // 解析txt文件内容，创建面要素
                    foreach (var path in list_txtPath)
                    {
                        string shp_name = path[(path.LastIndexOf("\\") + 1)..];

                        // 错误信息
                        string errMessage = $"【{shp_name}】\r";

                        // 获取txt文件的文本内容
                        string text = TxtTool.GetTXTContent(path);

                        // 文本中的【@】符号放前
                        string updata_text = ChangeSymbol(text);

                        // 获取属性描述
                        Dictionary<string, string> dict = GetAtt(text);
                        // 属性描述检查
                        errMessage += CheckAtt(dict);

                        // 获取坐标点文本
                        string[] fcs_text = updata_text.Split("@");
                        // 去除第一部分非坐标文本
                        List<string> fcs_text2List = new List<string>(fcs_text);
                        fcs_text2List.RemoveAt(0);

                        // 一个文件可能有多要素
                        foreach (var txt in fcs_text2List)
                        {
                            // 获取要素的部件数
                            int parts = GetCount(txt);

                            // 构建坐标点集合
                            var vertices_list = new List<List<Coordinate2D>>();
                            for (int i = 0; i < parts; i++)
                            {
                                var vertices = new List<Coordinate2D>();
                                vertices_list.Add(vertices);
                            }

                            // 编号、名称、类型、图幅、用途
                            int JZDS = 0;
                            double DKMJ = 0;
                            string DKBH = "";
                            string DKMC = "";
                            string JLTXS = "";
                            string TFH = "";
                            string DKYT = "";
                            string DLBM = "";
                            // 根据换行符分解坐标点文本
                            string[] list_point = txt.Split("\n");

                            // 监视部件号变化
                            string partID = "-1";
                            int pID = -1;

                            // 收集点行
                            List<string> pointTexts = new List<string>();

                            foreach (string pt in list_point)
                            {
                                // 清理一下
                                string point = pt.Replace("\r", "").Replace("\n", "");

                                // 如果存在中文的逗号
                                if (point.Contains('，'))
                                {
                                    errMessage += $"当前行[{point}]：存在中文的逗号\r";
                                }
                                // 如果存在空格
                                if (point.Contains(' '))
                                {
                                    errMessage += $"当前行[{point}]：存在空格\r";
                                }
                                // 如果点行缺少项
                                int symCount = TxtTool.StringInCount(point, ",");
                                if (symCount < 3 && symCount > 0)
                                {
                                    if (point[0] == 'J' || point[0] == 'G' || point[0].ToString().ToInt() != 0)   // 点行
                                    {
                                        errMessage += $"当前行[{point}]：缺少点号、圈号或XY坐标\r";
                                    }
                                }

                                if (TxtTool.StringInCount(point, ",") == 8)     // 名称、地块编号、功能文本
                                {
                                    JZDS = point.Split(",")[0].ToInt();
                                    DKMJ = point.Split(",")[1].ToDouble();
                                    DKBH = point.Split(",")[2];
                                    DKMC = point.Split(",")[3];
                                    JLTXS = point.Split(",")[4];
                                    TFH = point.Split(",")[5];
                                    DKYT = point.Split(",")[6];
                                    DLBM = point.Split(",")[7];
                                }

                                else if (TxtTool.StringInCount(point, ",") == 3)           // 点坐标文本
                                {
                                    // 点行收集
                                    pointTexts.Add(point);
                                    // 检查点行
                                    errMessage += CheckPointAtt(point);

                                    string fid = point.Split(",")[1].Replace(" ", "");        // 图斑部件号
                                    if (fid != partID)
                                    {
                                        pID += 1;
                                        partID = fid;
                                    }
                                    double lat = double.Parse(point.Split(",")[3].Replace(" ", ""));         // 经度
                                    double lng = double.Parse(point.Split(",")[2].Replace(" ", ""));         // 纬度

                                    vertices_list[pID].Add(new Coordinate2D(lat, lng));    // 加入坐标点集合
                                }
                                else     // 跳过无坐标部份的文本
                                {
                                    continue;
                                }
                            }

                            // 点行集合检查
                            errMessage += CheckPointList(pointTexts, isReverse);

                        }

                        // 错误提示汇总
                        errResult += errMessage;
                    }


                    // 复制模板
                    DirTool.CopyResourceFile(@$"CCTool.Data.Word.TXT质检结果.docx", wordPath);
                    // 输出检查结果
                    WordTool.WordRepalceText(wordPath, "{TXT检查}", errResult);

                    pw.AddMessageMiddle(0, errResult, Brushes.Blue);

                    pw.AddMessageEnd();
                });
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }


        private void openSHPButton_Click(object sender, RoutedEventArgs e)
        {
            // 打开TXT文件夹
            textFolderPath.Text = UITool.OpenDialogFolder();
            // 更新列表框的内容
            UpdataListBox();
        }

        // 更新列表框的内容
        private void UpdataListBox()
        {
            string folder = textFolderPath.Text;
            // 清除listbox
            listbox_txt.Items.Clear();
            // 生成TXT要素列表
            if (folder != "" && Directory.Exists(folder))
            {
                // 获取所有shp文件
                var files = DirTool.GetAllFiles(folder, ".txt");
                foreach (var file in files)
                {
                    // 将txt文件做成checkbox放入列表中
                    CheckBox cb = new CheckBox();
                    cb.Content = file.Replace(folder, "");
                    cb.IsChecked = true;
                    listbox_txt.Items.Add(cb);
                }
            }
        }

        // 文本中的【@】符号放前
        public static string ChangeSymbol(string text)
        {
            string[] lins = text.Split('\n');
            string updata_lins = "";
            foreach (string line in lins)
            {

                if (line.Contains("@"))
                {
                    string newline = line.Replace("@", "");
                    newline = "@" + newline;
                    updata_lins += newline + "\n";
                }
                else
                {
                    updata_lins += line + "\n";
                }
            }
            return updata_lins;
        }

        // 获取指标
        public static Dictionary<string, string> GetAtt(string text)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            // 如果不在标准段落里，返回空
            if (!text.Contains("[属性描述]") || !text.Contains("[地块坐标]"))
            {
                return dict;
            }
            else
            {
                string new_text = text[text.IndexOf("[属性描述]")..text.IndexOf("[地块坐标]")];
                string[] lines = new_text.Split("\n");
                foreach (string line in lines)
                {
                    if (line.Contains('='))
                    {
                        string before = line[..line.IndexOf("=")];
                        string after = line[(line.IndexOf("=") + 1)..];
                        dict.Add(before, after);
                    }
                }
                return dict;
            }
        }

        // 获取要素的部件数
        public static int GetCount(string lines)
        {
            List<string> indexs = new List<string>();

            // 根据换行符分解坐标点文本
            string[] list_point = lines.Split("\n");

            foreach (var point in list_point)
            {
                if (TxtTool.StringInCount(point, ",") == 3)          // 点坐标文本
                {
                    // 判断是否带空洞
                    string fid = point.Split(",")[1];        // 图斑部件号
                    if (!indexs.Contains(fid))
                    {
                        indexs.Add(fid);
                    }
                }
                else    // 路过非点坐标文本
                {
                    continue;
                }
            }

            return indexs.Count;
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "";
            UITool.Link2Web(url);
        }

        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_txt);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_txt);
        }

        // 属性行判断
        private string CheckAtt(Dictionary<string, string> dict)
        {
            string errMessage = "";

            // 必要的属性描述
            List<string> atts = new List<string>()
            {
                  "格式版本号","数据产生单位","数据产生日期","坐标系","几度分带",
                  "投影类型","计量单位","带号","精度","转换参数",
            };

            // 如果没有
            if (dict.Count == 0)
            {
                errMessage += "[属性描述]缺失\r";
            }
            else
            {
                errMessage += "[属性描述]\r";
                foreach (string att in atts)
                {
                    // 检查是否存在
                    if (!dict.ContainsKey(att))
                    {
                        errMessage += $"缺少属性行：[{att}]\r";
                    }
                    else
                    {
                        // 存在但值为空
                        if (dict[att] == "")
                        {
                            errMessage += $"[{att}]的属性值为空\r";
                        }
                    }
                }
            }

            return errMessage;
        }


        // 点行判断
        private string CheckPointAtt(string pt)
        {
            string errMessage = "";

            // 清一下空格
            string point = pt.Replace(" ", "");

            // 点号
            string pid = point.Split(",")[0];
            if (pid == "")
            {
                errMessage += $"当前行[{point}]：圈号为空\r";
            }
            // 点号合规
            else
            {
                string firstStr = pid[0].ToString();   // 第一个文字
                string nextstr = pid[1..];   // 剩余的文字

                if (pid[0] == 'G' || pid[0] == 'J')
                {
                    if (nextstr.ToInt() == 0)
                    {
                        errMessage += $"当前行[{point}]：点号有误\r";
                    }
                }
                else
                {
                    if (pid.ToInt() == 0)
                    {
                        errMessage += $"当前行[{point}]：点号有误\r";
                    }
                }
            }

            // 圈号
            string fid = point.Split(",")[1];
            if (fid == "")
            {
                errMessage += $"当前行[{point}]：圈号为空\r";
            }
            // 不是数字的话
            if (fid.ToInt() == 0)
            {
                errMessage += $"当前行[{point}]：圈号不是数字\r";
            }

            // x轴
            string xx = point.Split(",")[2];
            if (xx == "")
            {
                errMessage += $"当前行[{point}]：X坐标为空\r";
            }
            else
            {
                // 先检查有没有小数点
                if (!xx.Contains('.')) { xx += ".00"; }

                int count = xx[..xx.IndexOf('.')].Length;    // 整数位数
                if (count != 7)
                {
                    errMessage += $"当前行[{point}]：X坐标整数位数不是7位\r";
                }
            }

            // y轴
            string yy = point.Split(",")[3];
            if (yy == "")
            {
                errMessage += $"当前行[{point}]：X坐标为空\r";
            }
            else
            {
                // 先检查有没有小数点
                if (!yy.Contains('.')) { yy += ".00"; }

                int count = yy[..yy.IndexOf('.')].Length;    // 整数位数
                if (count != 8)
                {
                    errMessage += $"当前行[{point}]：Y坐标整数位数不是8位\r";
                }
            }

            return errMessage;
        }


        private string CheckPointList(List<string> pointTexts, bool isReverse)
        {
            string errMessage = "";

            if (pointTexts.Count < 3)
            {
                errMessage += $"点坐标少于3对\r";
            }
            // 首末点检查
            string firstRow = pointTexts[0].Replace("\r", "").Replace("\n", "");
            string lastRow = pointTexts[pointTexts.Count - 1].Replace("\r", "").Replace("\n", "");

            // 点号
            string firstP = firstRow.Split(",")[0];
            string lastP = lastRow.Split(",")[0];

            // 不勾选，按返回的情况
            if (isReverse == false)
            {
                if (firstP != lastP)
                {
                    errMessage += $"最后一个坐标点没有按要求返回\r";
                }
                // 返回的点行，检查是否一致
                else
                {
                    if (firstRow != lastRow)
                    {
                        errMessage += $"最后一个坐标点和第一个点不一致\r";
                    }
                }
            }
            // 不返回的情况，检查是否一致
            else
            {
                if (firstP == lastP)
                {
                    errMessage += $"最后一个坐标点不允许返回\r";
                }
            }

            return errMessage;
        }

        private void openWordButton_Click(object sender, RoutedEventArgs e)
        {
            textWordPath.Text = UITool.SaveDialogWord();
        }
    }
}
