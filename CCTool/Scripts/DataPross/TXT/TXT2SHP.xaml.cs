using ActiproSoftware.Windows.Extensions;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Data.Exceptions;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.GeoProcessing;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells.Drawing;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Windows;
using NPOI.OpenXmlFormats.Shared;
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
using SpatialReference = ArcGIS.Core.Geometry.SpatialReference;

namespace CCTool.Scripts.DataPross.TXT
{
    /// <summary>
    /// Interaction logic for TXT2SHP.xaml
    /// </summary>
    public partial class TXT2SHP : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "TXT2SHP";

        public TXT2SHP()
        {
            InitializeComponent();

            EventCenter.AddListener(EventDefine.UpdataCod, UpdataCod);

            // 初始化
            textFolderPath.Text = BaseTool.ReadValueFromReg(toolSet, "folderPath");
            textFeatureClassPath.Text = BaseTool.ReadValueFromReg(toolSet, "shpPath");
            textCod.Text = BaseTool.ReadValueFromReg(toolSet, "srName");

            // 更新列表框
            UpdataListBox();
        }

        // 更新坐标系信息
        private void UpdataCod()
        {
            textCod.Text = GlobalData.sr.Name;
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "TXT文件转SHP";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string folderPath = textFolderPath.Text;
                string shpPath = textFeatureClassPath.Text;
                string srName = textCod.Text;

                var cb_txts = listbox_txt.Items;

                // 参数写入本地
                BaseTool.WriteValueToReg(toolSet, "folderPath", folderPath);
                BaseTool.WriteValueToReg(toolSet, "shpPath", shpPath);
                BaseTool.WriteValueToReg(toolSet, "srName", srName);


                // 判断参数是否选择完全
                if (folderPath == "" || shpPath == "" || srName == "" || cb_txts.Count == 0)
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

                // 如果不存在目标文件夹，就创建一个
                if (!Directory.Exists(shpPath))
                {
                    Directory.CreateDirectory(shpPath);
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                // 异步执行
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(list_txtPath);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    foreach (string txtPath in list_txtPath)
                    {
                        string shp_name_old = txtPath[(txtPath.LastIndexOf(@"\") + 1)..].Replace(".txt", "");  // 获取要素名
                        // 预处理一下要素名，避免一些奇奇怪怪的符号
                        string shp_name = shp_name_old.Replace(".", "_");

                        pw.AddMessageMiddle(10, $@"创建要素：{shp_name}");

                        // 创建一个空要素
                        Arcpy.CreateFeatureclass(shpPath, shp_name, "POLYGON", srName);
                        // 新建字段
                        Arcpy.AddField(@$"{shpPath}\{shp_name}.shp", "文件名", "TEXT");
                        Arcpy.AddField(@$"{shpPath}\{shp_name}.shp", "编号", "TEXT");
                        Arcpy.AddField(@$"{shpPath}\{shp_name}.shp", "名称", "TEXT");
                        Arcpy.AddField(@$"{shpPath}\{shp_name}.shp", "类型", "TEXT");
                        Arcpy.AddField(@$"{shpPath}\{shp_name}.shp", "图幅", "TEXT");
                        Arcpy.AddField(@$"{shpPath}\{shp_name}.shp", "用途", "TEXT");

                        // 打开shp
                        FileSystemConnectionPath fileConnection = new FileSystemConnectionPath(new Uri(shpPath), FileSystemDatastoreType.Shapefile);
                        using FileSystemDatastore shapefile = new FileSystemDatastore(fileConnection);
                        // 获取FeatureClass
                        FeatureClass featureClass = shapefile.OpenDataset<FeatureClass>(shp_name);

                        // 获取txt文件的文本内容
                        string text = TxtTool.GetTXTContent(txtPath);

                        // 文本中的【@】符号放前
                        string updata_text = ChangeSymbol(text);

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
                            string bh = "";
                            string mc = "";
                            string lx = "";
                            string tf = "";
                            string yt = "";
                            // 根据换行符分解坐标点文本
                            string[] list_point = txt.Split("\n");

                            // 监视部件号变化
                            string partID = "-1";
                            int pID = -1;

                            foreach (var point in list_point)
                            {
                                if (TxtTool.StringInCount(point, ",") == 8)     // 名称、地块编号、功能文本
                                {
                                    bh = point.Split(",")[2];
                                    mc = point.Split(",")[3];
                                    lx = point.Split(",")[4];
                                    tf = point.Split(",")[5];
                                    yt = point.Split(",")[6];
                                }
                                else if (TxtTool.StringInCount(point, ",") == 3)           // 点坐标文本
                                {
                                    string fid = point.Split(",")[1].Replace(" ", "");        // 图斑部件号
                                    if (fid != partID)
                                    {
                                        pID += 1;
                                        partID= fid;
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
                                rowBuffer["文件名"] = shp_name.Replace(".txt", "");
                                rowBuffer["编号"] = bh;
                                rowBuffer["名称"] = mc;
                                rowBuffer["类型"] = lx;
                                rowBuffer["图幅"] = tf;
                                rowBuffer["用途"] = yt;

                                PolygonBuilderEx pb = new PolygonBuilderEx(vertices_list[0]);
                                // 如果有空洞，则添加内部Polygon
                                if (vertices_list.Count > 1)
                                {
                                    for (int i = 0; i < vertices_list.Count - 1; i++)
                                    {
                                        pb.AddPart(vertices_list[i + 1]);
                                    }
                                }
                                // 给新添加的行设置形状
                                rowBuffer[featureClassDefinition.GetShapeField()] = pb.ToGeometry();

                                // 在表中创建新行
                                using Feature feature = featureClass.CreateRow(rowBuffer);
                                context.Invalidate(feature);      // 标记行为无效状态
                            }, featureClass);

                            // 执行编辑操作
                            editOperation.Execute();
                        }
                        // 保存编辑
                        Project.Current.SaveEditsAsync();
                    }

                    pw.AddMessageEnd();
                });
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        private void openFeatureClassButton_Click(object sender, RoutedEventArgs e)
        {
            textFeatureClassPath.Text = UITool.OpenDialogFolder();
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
            string new_text = text[text.IndexOf("[项目信息]")..text.IndexOf("[属性描述]")];
            string[] lines = new_text.Split("\n");
            foreach (string line in lines)
            {
                if (line.Contains("="))
                {
                    string before = line[..line.IndexOf("=")];
                    string after = line[(line.IndexOf("=") + 1)..];
                    dict.Add(before, after);
                }
            }
            return dict;
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

        // 获取要素数据
        public static List<string> GetParts(string lines)
        {
            // 获取坐标点文本
            string[] fcs_text = lines.Split("@");

            // 去除第一部分非坐标文本
            List<string> fcs_text2List = new List<string>(fcs_text);
            fcs_text2List.RemoveAt(0);

            List<string> indexs = new List<string>();
            // 获取要素个数
            foreach (var line in fcs_text2List)
            {
                // 根据换行符分解坐标点文本
                string[] list_point = line.Split("\r");
                string fid = list_point[3].Split(",")[1];        // 图斑部件号
                if (!indexs.Contains(fid))
                {
                    indexs.Add(fid);
                }
            }
            // 生成列表
            List<string> parts = new List<string>();
            for (int i = 0; i < indexs.Count; i++)
            {
                parts.Add("");
            }

            foreach (var line in fcs_text2List)
            {
                // 根据换行符分解坐标点文本
                string[] list_point = line.Split("\r");
                int fid = int.Parse(list_point[3].Split(",")[1]);        // 图斑部件号
                parts[fid - 1] += "@";
                parts[fid - 1] += line;
            }
            return parts;
        }

        private void downLoadButton_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://pan.baidu.com/s/1KsCdwfIMOS61G6kHHO_SgQ?pwd=9kri";
            UITool.Link2Web(url);
        }

        private List<string> CheckData(List<string> in_data)
        {
            List<string> result = new List<string>();

            foreach (string txtPath in in_data)
            {
                string shp_name_old = txtPath[(txtPath.LastIndexOf(@"\") + 1)..].Replace(".txt", "");  // 获取要素名
                // 预处理一下要素名，避免一些奇奇怪怪的符号
                string shp_name = shp_name_old.Replace(".", "_");

                // 获取txt文件的文本内容
                string text = TxtTool.GetTXTContent(txtPath);
                // 文本中的【@】符号放前
                string updata_text = ChangeSymbol(text);

                // 获取坐标点文本
                string[] fcs_text = updata_text.Split("@");
                // 去除第一部分非坐标文本
                List<string> fcs_text2List = new List<string>(fcs_text);
                fcs_text2List.RemoveAt(0);

                // 一个文件可能有多要素
                foreach (var txt in fcs_text2List)
                {
                    // 根据换行符分解坐标点文本
                    string[] list_point = txt.Split("\n");

                    foreach (var point in list_point)
                    {
                        if (TxtTool.StringInCount(point, ",") > 3)     // 名称、地块编号、功能文本
                        {
                            if (TxtTool.StringInCount(point, ",") != 8)
                            {
                                result.Add($"【{shp_name}】错误行：{point}");
                            }
                        }
                    }
                }
            }

            return result;
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/140745686";
            UITool.Link2Web(url);
        }

        // 打开坐标系选择框
        private CoordinateSystemWindow coordinateSystemWindow = null;
        private void btn_cod_Click(object sender, RoutedEventArgs e)
        {
            UITool.OpenCoordinateSystemWindow(coordinateSystemWindow);
        }

        private void fm_Unloaded(object sender, RoutedEventArgs e)
        {
            EventCenter.RemoveListener(EventDefine.UpdataCod, UpdataCod);
        }

        private void btn_clearCod_Click(object sender, RoutedEventArgs e)
        {
            // 清除坐标系
            textCod.Clear();
        }

        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_txt);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_txt);
        }
    }
}
