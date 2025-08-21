using ArcGIS.Core.Data;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
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
using Geometry = ArcGIS.Core.Geometry.Geometry;
using Polygon = ArcGIS.Core.Geometry.Polygon;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for InsectSta.xaml
    /// </summary>
    public partial class InsectSta : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "InsectSta";

        // 字段表1
        List<string> fields = new List<string>()
        {
             "ID","SEC","XMMC","XZQMC","XZQDM","XMYDLX",
             "PZNR","PFWH","PFSJ","年份",
             "PZZMJ","XZJSYDMJ","NZYMJ","GDMJ","WLYDMJ","ZSTDMJ",
        };

        public InsectSta()
        {
            InitializeComponent();

            // 初始化参数选项
            textExcelPath.Text = BaseTool.ReadValueFromReg(toolSet, "excelPath");
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "用地压盖统计";

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.SaveDialogExcel();
        }

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取数据
                string fc = combox_fc.ComboxText();
                string excelPath = textExcelPath.Text;

                BaseTool.WriteValueToReg(toolSet, "excelPath", excelPath);


                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                // 判断参数是否选择完全
                if (fc == "" || excelPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }


                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");

                    // 检查数据
                    List<string> errs = CheckData(fc);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    // 复制excel
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.重复统计表.xlsx", excelPath);

                    // 获取工作薄、工作表
                    string excelFile = ExcelTool.GetPath(excelPath);
                    int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
                    // 打开工作薄
                    Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                    // 打开工作表
                    Worksheet sheet = wb.Worksheets[sheetIndex];
                    // 获取cells
                    Cells cells = sheet.Cells;

                    pw.AddMessageMiddle(10, "按年份排序_倒序");
                    // 按年份排序
                    string sort_fc = $@"{Project.Current.DefaultGeodatabasePath}\sort_fc";
                    Arcpy.Sort(fc, sort_fc, "年份 DESCENDING", "UR");

                    // 保存一个字典记录
                    List<List<string>> idDict = new List<List<string>>();

                    // 获取原始图层的要素类
                    FeatureClass featureClass = sort_fc.TargetFeatureClass();

                    // ID字段
                    string idField = sort_fc.TargetIDFieldName();

                    pw.AddMessageMiddle(20, "查找重叠图斑，提取信息");
                    // 获取目标图层和源图层的要素游标
                    using RowCursor originCursor = featureClass.Search();
                    // 初始行
                    int index = 3;

                    // 遍历源图层的要素
                    while (originCursor.MoveNext())
                    {
                        using Feature originFeature = (Feature)originCursor.Current;

                        string originID = originFeature[idField]?.ToString();

                        // 获取源要素的几何
                        Geometry originGeometry = originFeature.GetShape();

                        // 创建空间查询过滤器，以获取与源要素有重叠的目标要素
                        SpatialQueryFilter spatialFilter = new SpatialQueryFilter
                        {
                            FilterGeometry = originGeometry,
                            SpatialRelationship = SpatialRelationship.Intersects
                        };

                        // 原始要素标记
                        int originIndex = 0;

                        // 在目标图层中查询与源要素重叠的要素
                        using RowCursor identityCursor = featureClass.Search(spatialFilter);

                        // 记录lx和bz
                        List<string> lx_all = new List<string>();
                        string bz_result = "";
                        double bz_mj = 0;

                        while (identityCursor.MoveNext())
                        {
                            using Feature identityFeature = (Feature)identityCursor.Current;

                            string identityID = identityFeature[idField]?.ToString();

                            // 获取目标要素的几何
                            Geometry identityGeometry = identityFeature.GetShape();

                            // 源要素与目标要素的重叠
                            Polygon intersection = GeometryEngine.Instance.Intersection(originGeometry, identityGeometry) as Polygon;


                            // 如果不相交，或者是同一个图斑，就返回
                            if (intersection is null || originID == identityID || isInDict(idDict, originID, identityID))
                            {
                                continue;
                            }

                            // 相交面积
                            double mj = Math.Round(intersection.Area, 0);

                            if (mj == 0)
                            {
                                continue;
                            }

                            // 获取相交要素的字段值
                            string ID = identityFeature["ID"]?.ToString();
                            string SEC = identityFeature["SEC"]?.ToString();
                            string PZNR = identityFeature["PZNR"]?.ToString();
                            string PFWH = identityFeature["PFWH"]?.ToString();

                            string PWSJLY = identityFeature["PWSJLY"]?.ToString();

                            string originBS = originFeature["XMYDLX"]?.ToString();
                            string identityBS = identityFeature["XMYDLX"]?.ToString();

                            string lx = LX(originBS, identityBS);
                            string bz = $"压占{PFWH}（{PWSJLY[..1]}）-{lx}-{mj}平方米";
                            // 记录
                            idDict.Add(new List<string>() { originID, identityID });

                            // 如果是第一次
                            if (originIndex == 0)
                            {
                                // 复制行
                                if (index > 3)
                                {
                                    cells.CopyRow(cells, 3, index);
                                }

                                // 获取原始要素的字段值
                                for (int i = 0; i < fields.Count; i++)
                                {
                                    // 写入字段值
                                    cells[index, i + 2].Value = originFeature[fields[i]]?.ToString();
                                }
                                // 写入序号
                                cells[index, 0].Value = (index - 2).ToString();
                                cells[index, 1].Value = "重庆市";

                                // 写入
                                cells[index, 20].Value = ID;
                                cells[index, 21].Value = SEC;
                                cells[index, 22].Value = PZNR;

                                // 记录lx和bz
                                lx_all.Add(lx);
                                bz_result = bz;
                                bz_mj = mj;

                                cells[index, 18].Value = lx;
                                cells[index, 23].Value = bz_result;

                            }
                            // 如果不是第一次
                            else
                            {
                                index--;
                                // 记录lx和bz
                                if (!lx_all.Contains(lx)) { lx_all.Add(lx); }
                                if (mj>bz_mj)
                                {
                                    bz_result = bz;
                                    bz_mj = mj;
                                }
                                // 更新
                                string lx_result = "";
                                foreach (string item in lx_all)
                                {
                                    lx_result += (item + "\n");
                                }
                                cells[index, 18].Value = lx_result[..^1];
                                cells[index, 23].Value = bz_result;

                            }
                            index++;
                            originIndex++;
                        }
                    }

                    // 保存
                    wb.Save(excelFile);
                    wb.Dispose();

                    // 合并SEC列
                    ExcelTool.MergeSameCol(excelPath, 21);


                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        // 判断是否在字典里，反向也是
        private bool isInDict(List<List<string>> dict, string key, string value)
        {
            bool bo = false;

            foreach (var item in dict)
            {
                if (item[0] == value && item[1] == key)
                {
                    bo = true;
                }
            }
            return bo;
        }

        // 判断标识类型
        private string LX(string originBS_a, string identityBS_a)
        {
            string result = "xxxxxx";

            string originBS = originBS_a.Contains("国务院批准") ? "城镇村批次用地" : originBS_a;
            string identityBS = identityBS_a.Contains("国务院批准") ? "城镇村批次用地" : identityBS_a;

            if (originBS.Contains("单独选址") && identityBS.Contains("城镇村批次"))
            {
                result = "010300";
            }
            else if (originBS.Contains("城镇村批次") && identityBS.Contains("单独选址"))
            {
                result = "010300";
            }
            else if (originBS.Contains("单独选址") && identityBS.Contains("单独选址"))
            {
                result = "010600";
            }
            else if (originBS.Contains("城镇村批次") && identityBS.Contains("城镇村批次"))
            {
                result = "010400";
            }
            return result;
        }


        private List<string> CheckData(string fc)
        {
            List<string> result = new List<string>();

            string fieldEmptyResult = CheckTool.IsHaveFieldInLayer(fc, fields);
            if (fieldEmptyResult != "")
            {
                result.Add(fieldEmptyResult);
            }


            string fieldEmptyResult2 = CheckTool.IsHaveFieldInLayer(fc, "年份");
            if (fieldEmptyResult2 != "")
            {
                result.Add(fieldEmptyResult2);
            }

            return result;
        }
    }

}

