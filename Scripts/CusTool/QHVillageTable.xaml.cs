using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.OpenXmlFormats.Vml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Brushes = System.Windows.Media.Brushes;

namespace CCTool.Scripts.CusTool
{
    /// <summary>
    /// Interaction logic for QHVillageTable.xaml
    /// </summary>
    public partial class QHVillageTable : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public QHVillageTable()
        {
            InitializeComponent();

            // 初始化combox
            combox_unit.Items.Add("平方米");
            combox_unit.Items.Add("公顷");
            combox_unit.Items.Add("平方公里");
            combox_unit.Items.Add("亩");
            combox_unit.SelectedIndex = 1;
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "青海省村规结构调整表";

        // 点击打开按钮，选择输出的Excel文件位置
        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            // 打开Excel文件
            string path = UITool.SaveDialogExcel();
            // 将Excel文件的路径置入【textExcelPath】
            textExcelPath.Text = path;
        }

        // 添加要素图层的所有字符串字段到combox中
        private void combox_field_DropDown(object sender, EventArgs e)
        {
            // 将图层字段加入到Combox列表中
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_bmField);
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        // 运行
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
                string areaField = combox_areaField.ComboxText();

                string fc_path_gh = combox_fc_gh.ComboxText();
                string field_bm_gh = combox_bmField_gh.ComboxText();
                string areaField_gh = combox_areaField_gh.ComboxText();

                string unit = combox_unit.Text;

                // 单位系数设置
                double unit_xs = unit switch
                {
                    "平方米"=>1,
                    "公顷" => 10000,
                    "平方公里" => 100000,
                    "亩" => 666.66667,
                    _ => 1,
                };

                string excel_path = textExcelPath.Text;

                string output_table = init_gdb + @"\output_table";

                string czcField = "CZCSXM";   //  城镇村属性码

                await QueuedTask.Run(() =>
                {
                    if (!GisTool.IsHaveFieldInTarget(fc_path, czcField))
                    {
                        MessageBox.Show("现状图层缺少【CZCSXM】字段！");
                    }
                    if (!GisTool.IsHaveFieldInTarget(fc_path_gh, czcField))
                    {
                        MessageBox.Show("规划图层缺少【CZCSXM】字段！");
                    }
                });

                // 判断参数是否选择完全
                if (fc_path == "" || field_bm == "" || areaField == "" || fc_path_gh == "" || field_bm_gh == "" || areaField_gh == "" || excel_path == "")
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
                    List<string> errs = CheckData(fc_path, fc_path_gh, field_bm, field_bm_gh);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10, "复制Excel文件");
                    // 汇总用地
                    string excelName = "【青海】空间功能结构调整表";
                    string excelMapper = "【青海】用地用海代码_村庄功能";

                    // 复制嵌入资源中的Excel文件
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.{excelName}.xlsx", excel_path);
                    string excel_sheet = excel_path + @"\Sheet1$";
                    // 复制映射表
                    DirTool.CopyResourceFile(@$"CCTool.Data.Excel.{excelMapper}.xlsx", @$"{init_foder}\{excelMapper}.xlsx");

                    pw.AddMessageMiddle(10, "村庄功能映射");
                    // 初步汇总
                    // 调用GP工具【汇总】
                    Arcpy.Statistics(fc_path, output_table, $"{areaField} SUM", $"{czcField};{field_bm}");
                    Arcpy.Statistics(fc_path_gh, output_table + "_gh", $"{areaField_gh} SUM", $"{czcField};{field_bm_gh}");

                    // 添加村庄功能字段
                    string gnField = "vgGN";
                    Arcpy.AddField(output_table, gnField, "TEXT");
                    Arcpy.AddField(output_table + "_gh", gnField, "TEXT");

                    // 村庄功能映射
                    string mapper = @$"{init_foder}\{excelMapper}.xlsx\sheet1$";
                    ComboTool.AttributeMapper(output_table, field_bm, gnField, mapper);
                    ComboTool.AttributeMapper(output_table + "_gh", field_bm_gh, gnField, mapper);
                    // 区分202、203等属性
                    Segment(output_table, field_bm, gnField);
                    Segment(output_table + "_gh", field_bm_gh, gnField);

                    pw.AddMessageMiddle(10, $"汇总指标", Brushes.Gray);
                    // 汇总指标
                    List<string> bmList = new List<string>() { gnField };
                    string total = "国土面积";
                    Dictionary<string, double> dic = ComboTool.MultiStatisticsToDic(output_table, $"Sum_{areaField}", bmList, total, unit_xs);
                    Dictionary<string, double> dic_gh = ComboTool.MultiStatisticsToDic(output_table + "_gh", $"Sum_{areaField}", bmList, total, unit_xs);

                    pw.AddMessageMiddle(10, $"写入Excel", Brushes.Gray);
                    // 属性映射现状用地
                    ExcelTool.AttributeMapperDouble(excel_sheet, 10, 4, dic, 5);
                    // 属性映射规划用地
                    ExcelTool.AttributeMapperDouble(excel_sheet, 10, 6, dic_gh, 5);

                    // 删除0值行
                    ExcelTool.DeleteNullRow(excel_sheet, new List<int>() { 4, 6 }, 5);
                    // 删除指定列
                    ExcelTool.DeleteCol(excel_sheet, new List<int>() { 10 });

                    // 更改单位格
                    ExcelTool.WriteCell(excel_sheet, 2,1, $"单位：{unit}、%");

                    pw.AddMessageMiddle(10, "删除中间数据");
                    Arcpy.Delect(output_table);
                    Arcpy.Delect(output_table + "_gh");

                    // 复制映射表
                    File.Delete(@$"{init_foder}\{excelMapper}.xlsx");
                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        // 区分202、203等属性
        public void Segment(string tablePath, string bmField, string gnField)
        {
            Table table = tablePath.TargetTable();
            using RowCursor rowCursor = table.Search();
            while (rowCursor.MoveNext())
            {
                using Row row = rowCursor.Current;
                // 获取value
                string bm = row[bmField]?.ToString() ?? "";
                string gn = row[gnField]?.ToString() ?? "";
                string czcsxm = row["CZCSXM"]?.ToString() ?? "";
                // 村庄用地
                List<string> vgyd= new List<string>()
                {
                    "居住用地","公共管理与公共服务用地","商业服务业用地","工业用地","仓储用地","城镇村道路用地","交通场站用地","其他交通设施用地","公用设施用地","绿地与开敞空间用地","留白用地","空闲地",
                };
                // 城镇用地
                List<string> cityyd = new List<string>()
                {
                    "居住用地","公共管理与公共服务用地","商业服务业用地","工业用地","仓储用地","城镇村道路用地","交通场站用地","其他交通设施用地","公用设施用地","绿地与开敞空间用地","空闲地",
                };

                // 判断, 如果是城镇
                if (czcsxm == "201" || czcsxm == "202" || czcsxm == "201A" || czcsxm == "202A")
                {
                    if (cityyd.Contains(gn))
                    {
                        row[gnField] = "城镇" + gn;
                    }
                    else
                    {
                        row[gnField] = "城镇其他用地";
                    }
                }
                // 判断, 如果是村庄
                if (czcsxm == "203" || czcsxm == "203A")
                {
                    if (vgyd.Contains(gn))
                    {
                        row[gnField] = "村庄" + gn;
                    }
                    else
                    {
                        row[gnField] = "村庄其他用地";
                    }
                }

                row.Store();
            }
        }



        private void combox_areaField_DropDown(object sender, EventArgs e)
        {
            string fc_path = combox_fc.ComboxText();
            UITool.AddAllFloatFieldsToComboxPlus(fc_path, combox_areaField);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/138461488";
            UITool.Link2Web(url);
        }

        private void combox_fc_gh_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc_gh);
        }

        private void combox_bmField_gh_DropDown(object sender, EventArgs e)
        {
            // 将图层字段加入到Combox列表中
            UITool.AddTextFieldsToComboxPlus(combox_fc_gh.ComboxText(), combox_bmField_gh);
        }

        private void combox_areaField_gh_DropDown(object sender, EventArgs e)
        {
            string fc_path = combox_fc_gh.ComboxText();
            UITool.AddAllFloatFieldsToComboxPlus(fc_path, combox_areaField_gh);
        }


        // 检查输入数据
        private List<string> CheckData(string fc_path, string fc_path_gh, string field_bm, string field_bm_gh)
        {
            // 用地用海表_新
            Dictionary<string, string> ydyh = new Dictionary<string, string>()
            {
                { "01", "耕地"},{ "0101", "水田"},{ "0102", "水浇地"},{ "0103", "旱地"},{ "02", "园地"},{ "0201", "果园"},
                { "0202", "茶园"},{ "0203", "橡胶园地"},{ "0204", "油料园地"},{ "0205", "其他园地"},{ "03", "林地"},
                { "0301", "乔木林地"},{ "0302", "竹林地"},{ "0303", "灌木林地"},{ "0304", "其他林地"},{ "04", "草地"},
                { "0401", "天然牧草地"},{ "0402", "人工牧草地"},{ "0403", "其他草地"},{ "05", "湿地"},{ "0501", "森林沼泽"},
                { "0502", "灌丛沼泽"},{ "0503", "沼泽草地"},{ "0504", "其他沼泽地"},{ "0505", "沿海滩涂"},{ "0506", "内陆滩涂"},
                { "0507", "红树林地"},{ "06", "农业设施建设用地"},{ "060101", "村道用地"},{ "060102", "田间道"},
                { "0602", "设施农用地"},{ "060201", "种植设施建设用地"},{ "060202", "畜禽养殖设施建设用地"},{ "060203", "水产养殖设施建设用地"},
                { "07", "居住用地"},{ "0701", "城镇住宅用地"},{ "070101", "一类城镇住宅用地"},{ "070102", "二类城镇住宅用地"},
                { "070103", "三类城镇住宅用地"},{ "0702", "城镇社区服务设施用地"},{ "0703", "农村宅基地"},{ "070301", "一类农村宅基地"},
                { "070302", "二类农村宅基地"},{ "0704", "农村社区服务设施用地"},{ "08", "公共管理与公共服务用地"},{ "0801", "机关团体用地"},
                { "0802", "科研用地"},{ "0803", "文化用地"},{ "080301", "图书与展览用地"},{ "080302", "文化活动用地"},{ "0804", "教育用地"},
                { "080401", "高等教育用地"},{ "080402", "中等职业教育用地"},{ "080403", "中小学用地"},{ "080404", "幼儿园用地"},
                { "080405", "其他教育用地"},{ "0805", "体育用地"},{ "080501", "体育场馆用地"},{ "080502", "体育训练用地"},{ "0806", "医疗卫生用地"},
                { "080601", "医院用地"},{ "080602", "基层医疗卫生设施用地"},{ "080603", "公共卫生用地"},{ "0807", "社会福利用地"},
                { "080701", "老年人社会福利用地"},{ "080702", "儿童社会福利用地"},{ "080703", "残疾人社会福利用地"},{ "080704", "其他社会福利用地"},
                { "09", "商业服务业用地"},{ "0901", "商业用地"},{ "090101", "零售商业用地"},{ "090102", "批发市场用地"},{ "090103", "餐饮用地"},
                { "090104", "旅馆用地"},{ "090105", "公用设施营业网点用地"},{ "0902", "商务金融用地"},{ "0903", "娱乐用地"},
                { "0904", "其他商业服务业用地"},{ "1001", "工业用地"},{ "100101", "一类工业用地"},{ "100102", "二类工业用地"},
                { "100103", "三类工业用地"},{ "1002", "采矿用地"},{ "1003", "盐田"},{ "11", "仓储用地"},{ "1101", "物流仓储用地"},
                { "110101", "一类物流仓储用地"},{ "110102", "二类物流仓储用地"},{ "110103", "三类物流仓储用地"},{ "1102", "储备库用地"},
                { "1201", "铁路用地"},{ "1202", "公路用地"},{ "1203", "机场用地"},{ "1204", "港口码头用地"},
                { "1205", "管道运输用地"},{ "1206", "城市轨道交通用地"},{ "1207", "城镇村道路用地"},{ "1208", "交通场站用地"},
                { "120801", "对外交通场站用地"},{ "120802", "公共交通场站用地"},{ "120803", "社会停车场用地"},{ "1209", "其他交通设施用地"},
                { "1301", "供水用地"},{ "1302", "排水用地"},{ "1303", "供电用地"},{ "1304", "供燃气用地"},{ "1305", "供热用地"},
                { "1306", "通信用地"},{ "1307", "邮政用地"},{ "1308", "广播电视设施用地"},{ "1309", "环卫用地"},{ "1310", "消防用地"},
                { "1311", "水工设施用地"},{ "1312", "其他公用设施用地"},{ "14", "绿地与开敞空间用地"},{ "1401", "公园绿地"},{ "1402", "防护绿地"},
                { "1403", "广场用地"},{ "15", "特殊用地"},{ "1501", "军事设施用地"},{ "1502", "使领馆用地"},{ "1503", "宗教用地"},{ "1504", "文物古迹用地"},
                { "1505", "监教场所用地"},{ "1506", "殡葬用地"},{ "1507", "其他特殊用地"},{ "16", "留白用地"},{ "1701", "河流水面"},
                { "1702", "湖泊水面"},{ "1703", "水库水面"},{ "1704", "坑塘水面"},{ "1705", "沟渠"},{ "1706", "冰川及常年积雪"},{ "18", "渔业用海"},
                { "1801", "渔业基础设施用海"},{ "1802", "增养殖用海"},{ "1803", "捕捞海域"},{ "1804", "农林牧业用岛"},{ "19", "工矿通信用海"},
                { "1901", "工业用海"},{ "1902", "盐田用海"},{ "1903", "固体矿产用海"},{ "1904", "油气用海"},{ "1905", "可再生能源用海"},
                { "1906", "海底电缆管道用海"},{ "20", "交通运输用海"},{ "2001", "港口用海"},{ "2002", "航运用海"},{ "2003", "路桥隧道用海"},
                { "2004", "机场用海"},{ "2005", "其他交通运输用海"},{ "21", "游憩用海"},{ "2101", "风景旅游用海"},{ "2102", "文体休闲娱乐用海"},
                { "22", "特殊用海"},{ "2201", "军事用海"},{ "2202", "科研教育用海"},{ "2203", "海洋保护修复及海岸防护工程用海"},
                { "2204", "排污倾倒用海"},{ "2205", "水下文物保护用海"},{ "2206", "其他特殊用海"},{ "2301", "空闲地"},
                { "2302", "后备耕地"},{ "2303", "田坎"},{ "2304", "盐碱地"},{ "2305", "沙地"},{ "2306", "裸土地"},{ "2307", "裸岩石砾地"},{ "24", "其他海域"},
            };

            List<string> result = new List<string>();
            // 检查字段值是否符合要求
            string result_value = CheckTool.CheckFieldValue(fc_path, field_bm, ydyh.Keys.ToList());
            if (result_value != "")
            {
                result.Add(result_value);
            }

            string result_value2 = CheckTool.CheckFieldValue(fc_path_gh, field_bm_gh, ydyh.Keys.ToList());
            if (result_value2 != "")
            {
                result.Add(result_value2);
            }

            return result;
        }
    }
}
