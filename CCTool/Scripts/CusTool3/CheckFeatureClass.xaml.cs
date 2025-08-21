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
using NPOI.OpenXmlFormats.Dml.Diagram;
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
using SpatialReference = ArcGIS.Core.Geometry.SpatialReference;
using Table = ArcGIS.Core.Data.Table;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for CheckFeatureClass.xaml
    /// </summary>
    public partial class CheckFeatureClass : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "CheckFeatureClass";
        public CheckFeatureClass()
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
        string tool_name = "TXT与要素类一致性质检(吉)";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string folderPath = textFolderPath.Text;
                string fc = combox_fc.ComboxText();
                string wordPath = textWordPath.Text;

                var cb_txts = listbox_txt.Items;

                // 参数写入本地
                BaseTool.WriteValueToReg(toolSet, "folderPath", folderPath);
                BaseTool.WriteValueToReg(toolSet, "wordPath", wordPath);

                // 判断参数是否选择完全
                if (folderPath == "" || fc == "" || wordPath == "" || cb_txts.Count == 0)
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
                    pw.AddMessageStart("从TXT文件生成要素图层");

                    // 参数获取
                    string gdb_def = Project.Current.DefaultGeodatabasePath;
                    // 目标要素坐标系
                    FeatureLayer targetLayer = fc.TargetFeatureLayer();
                    SpatialReference sr = targetLayer.GetSpatialReference();

                    // 创建一个空要素
                    string tem_fc = "TXT导出的要素类";
                    Arcpy.CreateFeatureclass(gdb_def, tem_fc, "POLYGON", sr);
                    string targetFC = $@"{gdb_def}\{tem_fc}";

                    // 新建字段
                    GisTool.AddField(targetFC, "界址点数", FieldType.Integer);
                    GisTool.AddField(targetFC, "地块面积", FieldType.Double);
                    GisTool.AddField(targetFC, "地块编号");
                    GisTool.AddField(targetFC, "地块名称");
                    GisTool.AddField(targetFC, "记录图形属");
                    GisTool.AddField(targetFC, "图幅号");
                    GisTool.AddField(targetFC, "地块用途");
                    GisTool.AddField(targetFC, "地类编码");


                    // 打开数据库
                    using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_def))))
                    {
                        // 创建要素并添加到要素类中
                        using FeatureClass featureClass = gdb.OpenDataset<FeatureClass>(tem_fc);
                        // 解析txt文件内容，创建面要素
                        foreach (var path in list_txtPath)
                        {
                            string shp_name = path[(path.LastIndexOf("\\") + 1)..];

                            // 获取txt文件的文本内容
                            string text = TxtTool.GetTXTContent(path);

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

                                foreach (string point in list_point)
                                {
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
                                    rowBuffer["界址点数"] = JZDS;
                                    rowBuffer["地块面积"] = DKMJ;
                                    rowBuffer["地块编号"] = DKBH;
                                    rowBuffer["地块名称"] = DKMC;
                                    rowBuffer["记录图形属"] = JLTXS;
                                    rowBuffer["图幅号"] = TFH;
                                    rowBuffer["地块用途"] = DKYT;
                                    rowBuffer["地类编码"] = DLBM;

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

                        }
                    }
                    // 保存编辑
                    Project.Current.SaveEditsAsync();

                    // 将要素类添加到当前地图
                    MapCtlTool.AddLayerToMap(targetFC);

                    // 【检查生成的要素图层】
                    pw.AddMessageMiddle(30, "检查生成的要素图层");

                    // 错误提示汇总
                    string errResult = CheckGeometry(targetFC, fc);
                    // 复制模板
                    DirTool.CopyResourceFile(@$"CCTool.Data.Word.图形质检结果.docx", wordPath);
                    // 输出检查结果
                    WordTool.WordRepalceText(wordPath, "{图形检查}", errResult);

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

        private string CheckGeometry(string targetFC, string mapFC)
        {
            string errMessage = "";

            string def_gdb = Project.Current.DefaultGeodatabasePath;

            // 自相交等几何检查
            errMessage += "【自相交等几何检查】\r";

            string geoCheckResult = $@"{def_gdb}\geoCheckResult";
            Arcpy.CheckGeometry(targetFC, geoCheckResult, true);

            var dic_geoErr = GisTool.GetDictFromPath(geoCheckResult, "FEATURE_ID", "PROBLEM");

            if (dic_geoErr.Count != 0)
            {
                foreach (var geoErr in dic_geoErr)
                {
                    errMessage += $@"[OBJECTID_{geoErr.Key}]：{geoErr.Value}。" + "\r";
                }
            }
            // 面积是否有负值
            errMessage += "【面积是否有负值】\r";

            var dic_area = GisTool.GetDictFromPathDouble(targetFC, "OBJECTID", "Shape_Area");
            foreach (var area in dic_area)
            {
                if (area.Value <= 0)
                {
                    errMessage += $@"[OBJECTID_{area.Key}]：面积为负。" + "\r";
                }
            }
            // 自身重叠检查
            errMessage += "【自身重叠检查】\r";

            // 定义规则
            List<string> rules = new List<string>() { "Must Not Overlap (Area)" };
            ComboTool.TopologyCheck(targetFC, rules, def_gdb);
            // 获取错误数量
            string overlap_err = $@"{def_gdb}\TopErr_poly";
            var dic_topoErr = GisTool.GetDictFromPath(overlap_err, "OriginObjectID", "DestinationObjectID");

            if (dic_topoErr.Count != 0)
            {
                foreach (var dtopoErr in dic_topoErr)
                {
                    errMessage += $@"[OBJECTID_{dtopoErr.Key}和{dtopoErr.Value}]：存在重叠。" + "\r";
                }
                // 复制出来
                Arcpy.CopyFeatures(overlap_err, $@"{def_gdb}\自身重叠图斑");
            }

            // 和源数据检查
            errMessage += "【和源数据检查】\r";
            // 冗余地块
            string erase01 = $@"{def_gdb}\冗余地块";
            Arcpy.Erase(targetFC, mapFC, erase01);
            if (erase01.TargetTable().GetCount() > 0)
            {
                errMessage += "导出图斑存在冗余地块。\r";
                MapCtlTool.AddLayerToMap(erase01);
            }

            // 缺少地块
            string erase02 = $@"{def_gdb}\缺少地块";
            Arcpy.Erase(mapFC, targetFC, erase02);
            if (erase02.TargetTable().GetCount() > 0)
            {
                errMessage += "导出图斑缺少地块。\r";
                MapCtlTool.AddLayerToMap(erase02);
            }

            // 删除中间数据
            Arcpy.Delect(geoCheckResult);

            return errMessage;
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


        private void openWordButton_Click(object sender, RoutedEventArgs e)
        {
            textWordPath.Text = UITool.SaveDialogWord();
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }
    }
}
