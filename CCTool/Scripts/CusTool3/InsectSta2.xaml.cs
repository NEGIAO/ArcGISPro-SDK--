using ArcGIS.Core.Data;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using Microsoft.Office.Core;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using Org.BouncyCastle.Asn1.X509;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
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
using static QRCoder.PayloadGenerator;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Geometry = ArcGIS.Core.Geometry.Geometry;
using MessageBox = System.Windows.Forms.MessageBox;
using Polygon = ArcGIS.Core.Geometry.Polygon;
using SpatialReference = ArcGIS.Core.Geometry.SpatialReference;

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for InsectSta2.xaml
    /// </summary>
    public partial class InsectSta2 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "InsectSta2";

        public InsectSta2()
        {
            InitializeComponent();

            // 初始化参数选项
            textExcelPath.Text = BaseTool.ReadValueFromReg(toolSet, "excelFolder");

            // 预置
            UITool.InitFeatureLayerToComboxPlus(combox_pd, "用来叠加的征地批文库");
            UITool.InitFeatureLayerToComboxPlus(combox_gd, "2025年供地库");
            UITool.InitFeatureLayerToComboxPlus(combox_kfbj, "新CZKFBJGX");
            UITool.InitFeatureLayerToComboxPlus(combox_sthx, "生态保护红线STBHHX");
            UITool.InitFeatureLayerToComboxPlus(combox_jbnt, "永久基本农田YJJBNTBHTB");
            UITool.InitFeatureLayerToComboxPlus(combox_bgcc, "23变更调查擦除分析库");
            UITool.InitFeatureLayerToComboxPlus(combox_czc, "CZCDYD分析库");
            UITool.InitFeatureLayerToComboxPlus(combox_fq, "分区规划ydyh");

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "用地压盖统计2";

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.OpenDialogFolder();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取数据
                string pd = combox_pd.ComboxText();
                string gd = combox_gd.ComboxText();
                string kfbj = combox_kfbj.ComboxText();
                string sthx = combox_sthx.ComboxText();
                string jbnt = combox_jbnt.ComboxText();
                string bgcc = combox_bgcc.ComboxText();
                string czc = combox_czc.ComboxText();
                string fq = combox_fq.ComboxText();

                string excelFolder = textExcelPath.Text;

                BaseTool.WriteValueToReg(toolSet, "excelFolder", excelFolder);


                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                // 判断参数是否选择完全
                if (pd == "" || gd == "" || kfbj == "" || sthx == "" || jbnt == "" || bgcc == "" || czc == "" || fq == "" || excelFolder == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }


                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");

                    // 检查数据
                    List<string> errs = CheckData();
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
                    string excelPath1 = $@"{excelFolder}\表格一.xlsx";
                    string excelPath2 = $@"{excelFolder}\表格二.xlsx";

                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.杂七杂八.表格一.xlsx", excelPath1);
                    DirTool.CopyResourceFile(@"CCTool.Data.Excel.杂七杂八.表格二.xlsx", excelPath2);

                    pw.AddMessageMiddle(10, "创建表格一");

                    pw.AddMessageMiddle(10, "      分析批供地重叠情况", Brushes.Gray);
                    // 分析批供地重叠情况
                    CreateExcel10(pd, gd);

                    pw.AddMessageMiddle(10, "      写入表格一", Brushes.Gray);
                    // 写入表格一
                    CreateExcel11(excelPath1, pd, gd);

                    pw.AddMessageMiddle(10, "创建表格二");
                    // 生成唯一标记
                    string idField = "newID";
                    Arcpy.AddField(pd, idField, "LONG");
                    Arcpy.CalculateField(pd, idField, "SequentialNumber()", "rec=1\r\ndef SequentialNumber():\r\n    global rec\r\n    rec+=1\r\n    return rec");

                    // 生成已批未供用地
                    string defGDB = Project.Current.DefaultGeodatabasePath;
                    string pg = $@"{defGDB}\已批未供用地";
                    Arcpy.Erase(pd, gd, pg, true);

                    // 初始化批地属性库
                    List<pdAtt> pdAtts = new List<pdAtt>();
                    InitPdAtt(pdAtts, pd);

                    pw.AddMessageMiddle(10, "      分析三线压占情况", Brushes.Gray);
                    // 分析三线压占情况
                    CreateExcel20(pdAtts, pg, kfbj, "KFBJ", idField);
                    CreateExcel20(pdAtts, pg, sthx, "STHX", idField);
                    CreateExcel20(pdAtts, pg, jbnt, "YN", idField);

                    pw.AddMessageMiddle(10, "      分析城镇村", Brushes.Gray);
                    // 分析城镇村
                    CreateExcel21(pdAtts, pg, czc, idField);

                    pw.AddMessageMiddle(10, "      分析非建设用地", Brushes.Gray);
                    // 分析非建设用地
                    CreateExcel22(pdAtts, pg, bgcc, idField);

                    pw.AddMessageMiddle(10, "      分析规划用地", Brushes.Gray);
                    // 分析规划用地
                    CreateExcel23(pdAtts, pd, fq, idField);

                    pw.AddMessageMiddle(10, "      写入表格二", Brushes.Gray);
                    // 写入表格二
                    CreateExcel24(pdAtts, excelPath2, pd);

                    // 整理表格二
                    CreateExcel25(excelPath2);


                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void CreateExcel25(string excelPath)
        {
            List<int> delectIDs = new List<int>();

            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            Cells cells = sheet.Cells;
            // 逐行处理
            for (int i = cells.MaxDataRow; i >= 0; i--)
            {
                //  获取目标cell
                Cell lxCell = sheet.Cells[i, 18];
                // 属性映射
                if (lxCell is not null)
                {
                    string lx = lxCell.StringValue;
                    if (lx == "060100")
                    {
                        cells.DeleteRow(i);
                    }
                }
            }

            // 分解
            for (int i = cells.MaxDataRow; i >= 0; i--)
            {
                //  获取目标cell
                Cell flCell = sheet.Cells[i, 32];
                // 属性映射
                if (flCell is not null)
                {
                    string flmc = flCell.StringValue;

                    List<string> fls = flmc.Split(',').ToList();

                    for (int j = 0; j < fls.Count; j++)
                    {
                        cells.CopyRow(cells, i, i + j);
                        cells[i + j, 32].Value = fls[j];
                    }
                }
            }

            // 保存
            wb.Save(excelPath);
            wb.Dispose();

        }

        private void InitPdAtt(List<pdAtt> pdAtts, string pd)
        {
            long featureCount = pd.GetFeatureCount();
            for (int i = 0; i < featureCount; i++)
            {
                pdAtt pdAtt = new pdAtt();
                pdAtt.OID = (i + 1).ToString();
                pdAtts.Add(pdAtt);
            }
        }

        private void CreateExcel21(List<pdAtt> pdAtts, string origin, string czc, string idField)
        {
            // 获取原始图层和标识图层的要素类
            FeatureClass originFeatureClass = origin.TargetFeatureClass();
            FeatureClass targetFeatureClass = czc.TargetFeatureClass();

            // 获取目标图层和源图层的要素游标
            using RowCursor originCursor = originFeatureClass.Search();

            // 遍历源图层的要素，标记
            while (originCursor.MoveNext())
            {
                using Feature originFeature = (Feature)originCursor.Current;
                string OID = originFeature[idField].ToString();

                // 获取源要素的几何
                Polygon originGeometry = originFeature.GetShape() as Polygon;

                double cz_totalArea = 0;
                double ck_totalArea = 0;
                double qt_totalArea = 0;

                // 创建空间查询过滤器，以获取与源要素有重叠的目标要素
                SpatialQueryFilter spatialFilter = new SpatialQueryFilter
                {
                    FilterGeometry = originGeometry,
                    SpatialRelationship = SpatialRelationship.Intersects
                };

                // 在目标图层中查询与源要素重叠的要素
                using RowCursor targetCursor = targetFeatureClass.Search(spatialFilter);
                while (targetCursor.MoveNext())
                {
                    using Feature targetFeature = (Feature)targetCursor.Current;
                    // 获取目标要素的几何
                    Polygon targetGeometry = targetFeature.GetShape() as Polygon;

                    // 计算源要素与目标要素的重叠面积
                    Geometry intersection = GeometryEngine.Instance.Intersection(originGeometry, targetGeometry);

                    // 如果重叠面积大于当前最大重叠面积，则更新最大重叠面积和目标要素
                    if (intersection is not null)
                    {
                        double area = (intersection as Polygon).Area;
                        // 类型
                        string ydType = targetFeature["CZCLX"]?.ToString();
                        if (ydType == "203")
                        {
                            cz_totalArea += area;
                        }
                        else if (ydType == "204")
                        {
                            ck_totalArea += area;
                        }
                        else
                        {
                            qt_totalArea += area;
                        }

                    }
                }

                // 赋值
                foreach (var pdAtt in pdAtts)
                {
                    if (pdAtt.OID == OID)
                    {
                        pdAtt.BGCZMJ = cz_totalArea;
                        pdAtt.BGCKMJ = ck_totalArea;
                        pdAtt.BGQTJSMJ = qt_totalArea;
                        pdAtt.BGZMJ = originGeometry.Area;
                    }
                }
            }
        }

        private void CreateExcel22(List<pdAtt> pdAtts, string origin, string bgcc, string idField)
        {
            // 获取原始图层和标识图层的要素类
            FeatureClass originFeatureClass = origin.TargetFeatureClass();
            FeatureClass targetFeatureClass = bgcc.TargetFeatureClass();

            // 获取目标图层和源图层的要素游标
            using RowCursor originCursor = originFeatureClass.Search();

            // 遍历源图层的要素，标记
            while (originCursor.MoveNext())
            {
                using Feature originFeature = (Feature)originCursor.Current;
                string OID = originFeature[idField].ToString();

                // 获取源要素的几何
                Polygon originGeometry = originFeature.GetShape() as Polygon;

                double nyd_totalArea = 0;
                double gd_totalArea = 0;
                double wlyd_totalArea = 0;
                double qt = 0;

                // 创建空间查询过滤器，以获取与源要素有重叠的目标要素
                SpatialQueryFilter spatialFilter = new SpatialQueryFilter
                {
                    FilterGeometry = originGeometry,
                    SpatialRelationship = SpatialRelationship.Intersects
                };

                // 在目标图层中查询与源要素重叠的要素
                using RowCursor targetCursor = targetFeatureClass.Search(spatialFilter);
                while (targetCursor.MoveNext())
                {
                    using Feature targetFeature = (Feature)targetCursor.Current;
                    // 获取目标要素的几何
                    Polygon targetGeometry = targetFeature.GetShape() as Polygon;

                    // 计算源要素与目标要素的重叠面积
                    Geometry intersection = GeometryEngine.Instance.Intersection(originGeometry, targetGeometry);

                    // 如果重叠面积大于当前最大重叠面积，则更新最大重叠面积和目标要素
                    if (intersection is not null)
                    {
                        double area = (intersection as Polygon).Area;
                        // 类型
                        string ydType = targetFeature["DLBM"]?.ToString();

                        // 分类
                        List<string> nyds = new List<string>()
                        {
                            "0303","0304", "0306","0402",

                            "0101","0102", "0103",
                            "0201","0202", "0203","0204",
                            "0301","0302", "0305","0307",

                             "0401","0403", "0404",
                             "1006",
                             "1103","1104", "1107","1104A","1107A",
                             "1202","1203"
                        };
                        List<string> gds = new List<string>() { "0101", "0102", "0103" };
                        List<string> wlyds = new List<string>()
                        {
                            "1105", "1106", "1108",
                            "1101", "1102", "1110",
                            "1204", "1205", "1206","1207","1208",
                        };

                        if (nyds.Contains(ydType))
                        {
                            nyd_totalArea += area;
                        }
                        else if (gds.Contains(ydType))
                        {
                            gd_totalArea += area;
                        }
                        else if (wlyds.Contains(ydType))
                        {
                            wlyd_totalArea += area;
                        }
                        else
                        {
                            qt += area;
                        }

                    }
                }

                // 赋值
                foreach (var pdAtt in pdAtts)
                {
                    if (pdAtt.OID == OID)
                    {
                        pdAtt.BGNYDMJ = nyd_totalArea;
                        pdAtt.BGGDMJ = gd_totalArea;
                        pdAtt.BGWLYDMJ = wlyd_totalArea;

                        double area = pdAtt.BGQTJSMJ;
                        pdAtt.BGQTJSMJ = area + qt;
                    }
                }
            }


        }

        private void CreateExcel23(List<pdAtt> pdAtts, string origin, string fq, string idField)
        {
            // 添加标记字段
            string smField = "GHYDFLMC";
            if (!GisTool.IsHaveFieldInTarget(origin, smField))
            {
                Arcpy.AddField(origin, smField, "TEXT");
            }

            string defaultGDB = Project.Current.DefaultGeodatabasePath;

            // 获取原始图层和标识图层的要素类
            FeatureClass originFeatureClass = origin.TargetFeatureClass();
            FeatureClass targetFeatureClass = fq.TargetFeatureClass();

            FeatureLayer originFeatureLayer = origin.TargetFeatureLayer();
            FeatureLayer targetFeatureLayer = fq.TargetFeatureLayer();

            // 如果两个图层使用的空间参考不同，投影到一致
            SpatialReference sr_ori = originFeatureLayer.GetSpatialReference();
            SpatialReference sr_iden = targetFeatureLayer.GetSpatialReference();
            if (!sr_ori.Equals(sr_iden))
            {
                string MidFC = $@"{defaultGDB}\MidFC";
                Arcpy.Project(fq, $@"{defaultGDB}\MidFC", sr_ori.Wkt, sr_iden.Wkt);
                targetFeatureClass = MidFC.TargetFeatureClass();
            }


            // 获取目标图层和源图层的要素游标
            using RowCursor originCursor = originFeatureClass.Search();

            // 遍历源图层的要素，标记
            while (originCursor.MoveNext())
            {
                using Feature originFeature = (Feature)originCursor.Current;
                string OID = originFeature[idField].ToString();

                // 获取源要素的几何
                Polygon originGeometry = originFeature.GetShape() as Polygon;

                string sm = "";
                double pdPolygonArea = originGeometry.Area;

                // 创建空间查询过滤器，以获取与源要素有重叠的目标要素
                SpatialQueryFilter spatialFilter = new SpatialQueryFilter
                {
                    FilterGeometry = originGeometry,
                    SpatialRelationship = SpatialRelationship.Intersects
                };

                // 在目标图层中查询与源要素重叠的要素
                using RowCursor targetCursor = targetFeatureClass.Search(spatialFilter);
                while (targetCursor.MoveNext())
                {
                    using Feature targetFeature = (Feature)targetCursor.Current;
                    // 获取目标要素的几何
                    Polygon targetGeometry = targetFeature.GetShape() as Polygon;

                    // 计算源要素与目标要素的重叠面积
                    Geometry intersection = GeometryEngine.Instance.Intersection(originGeometry, targetGeometry);

                    // 如果重叠面积大于当前最大重叠面积，则更新最大重叠面积和目标要素
                    if (intersection is not null)
                    {
                        // 类型
                        string ydmc = targetFeature["GHYDFLDM"]?.ToString();
                        string fl = DMGL(ydmc);

                        if (!sm.Contains(fl))
                        {
                            sm += (fl + ",");
                        }
                    }
                }

                // 赋值
                foreach (var pdAtt in pdAtts)
                {
                    if (pdAtt.OID == OID)
                    {
                        if (sm != "")
                        {
                            pdAtt.GHYDFLMC = sm[..^1];
                        }
                    }
                }
            }

        }

        private string DMGL(string ydmc)
        {
            string fl = "";
            int dm = ydmc.ToInt();
            if (dm < 16 && dm > 5)
            {
                fl = "建设用地";
            }
            else if (dm == 16)
            {
                fl = "留白用地";
            }
            else
            {
                fl = "非建设用地";
            }

            return fl;
        }


        private void CreateExcel20(List<pdAtt> pdAtts, string origin, string target, string field, string idField)
        {

            // 获取原始图层和标识图层的要素类
            FeatureClass originFeatureClass = origin.TargetFeatureClass();
            FeatureClass targetFeatureClass = target.TargetFeatureClass();

            // 获取目标图层和源图层的要素游标
            using RowCursor originCursor = originFeatureClass.Search();

            // 遍历源图层的要素，标记
            while (originCursor.MoveNext())
            {
                using Feature originFeature = (Feature)originCursor.Current;
                string OID = originFeature[idField].ToString();

                // 获取源要素的几何
                Polygon originGeometry = originFeature.GetShape() as Polygon;

                double totalArea = 0;

                // 创建空间查询过滤器，以获取与源要素有重叠的目标要素
                SpatialQueryFilter spatialFilter = new SpatialQueryFilter
                {
                    FilterGeometry = originGeometry,
                    SpatialRelationship = SpatialRelationship.Intersects
                };

                // 在目标图层中查询与源要素重叠的要素
                using RowCursor targetCursor = targetFeatureClass.Search(spatialFilter);
                while (targetCursor.MoveNext())
                {
                    using Feature targetFeature = (Feature)targetCursor.Current;
                    // 获取目标要素的几何
                    Polygon targetGeometry = targetFeature.GetShape() as Polygon;

                    // 计算源要素与目标要素的重叠面积
                    Geometry intersection = GeometryEngine.Instance.Intersection(originGeometry, targetGeometry);

                    // 如果重叠面积大于当前最大重叠面积，则更新最大重叠面积和目标要素
                    if (intersection is not null)
                    {
                        double area = (intersection as Polygon).Area;
                        totalArea += area;
                    }
                }

                // 赋值
                foreach (var pdAtt in pdAtts)
                {
                    if (pdAtt.OID == OID)
                    {
                        if (field == "KFBJ")
                        {
                            pdAtt.ZYCZKFBJMJ = totalArea;
                        }
                        else if (field == "STHX")
                        {
                            pdAtt.ZYSTHXMJ = totalArea;
                        }
                        else if (field == "YN")
                        {
                            pdAtt.ZYYNMJ = totalArea;
                        }
                    }
                }
            }
        }

        private void CreateExcel24(List<pdAtt> pdAtts, string excelPath2, string pd)
        {
            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excelPath2);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath2);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取cells
            Cells cells = sheet.Cells;

            // 获取原始图层和标识图层的要素类
            FeatureClass pdFeatureClass = pd.TargetFeatureClass();

            int index = 7;
            int rowIndex = 0;
            using RowCursor pdCursor = pdFeatureClass.Search();
            // 遍历源图层的要素，写入表格
            while (pdCursor.MoveNext())
            {
                // 复制行
                cells.CopyRow(cells, 6, index);

                using Feature pdFeature = (Feature)pdCursor.Current;

                List<string> fds = new List<string>()
                {
                    "ID","SEC","XMMC","SZXZQDM","SZXZQMC","XMYDLX","PZNR","PFWH","PFSJ","PFSJ",
                    "PZZMJ","XZJSYDMJ","NZYMJ","GDMJ","WLYDMJ","ZSTDMJ",
                    "QLLX","BJBGYKNMJ","BNJSLYQX",
                    "ZYYNMJ","ZYSTHXMJ","ZYCZKFBJMJ","BGZMJ","BGNYDMJ",
                    "BGGDMJ","BGWLYDMJ","BGCZMJ","BGCKMJ","BGQTJSMJ",
                    "BZ","GHYDFLMC"

                };

                cells[index, 2].Value = pdFeature["ID"]?.ToString();
                cells[index, 3].Value = pdFeature["SEC"]?.ToString();
                cells[index, 4].Value = pdFeature["XMMC"]?.ToString();
                cells[index, 5].Value = pdFeature["SZXZQDM"]?.ToString();

                cells[index, 6].Value = pdFeature["SZXZQMC"]?.ToString();
                cells[index, 7].Value = pdFeature["XMYDLX"]?.ToString();
                cells[index, 8].Value = pdFeature["PZNR"]?.ToString();
                cells[index, 9].Value = pdFeature["PFWH"]?.ToString();
                cells[index, 10].Value = pdFeature["PFSJ"]?.ToString();

                cells[index, 12].Value = (double)pdFeature["PZZMJ"];
                cells[index, 13].Value = (double)pdFeature["XZJSYDMJ"];
                cells[index, 14].Value = (double)pdFeature["NZYMJ"];
                cells[index, 15].Value = (double)pdFeature["GDMJ"];
                cells[index, 16].Value = (double)pdFeature["WLYDMJ"];
                cells[index, 17].Value = (double)pdFeature["ZSTDMJ"];

                cells[index, 18].Value = pdFeature["QLLX"]?.ToString();
                cells[index, 19].Value = (double)pdFeature["BJBGYKNMJ"];
                cells[index, 20].Value = pdFeature["BNJSLYQX"]?.ToString();

                cells[index, 21].Value = Math.Round(pdAtts[rowIndex].ZYYNMJ / 10000, 4);
                cells[index, 22].Value = Math.Round(pdAtts[rowIndex].ZYSTHXMJ / 10000, 4);
                cells[index, 23].Value = Math.Round(pdAtts[rowIndex].ZYCZKFBJMJ / 10000, 4);
                cells[index, 24].Value = Math.Round(pdAtts[rowIndex].BGZMJ / 10000, 4);
                cells[index, 25].Value = Math.Round(pdAtts[rowIndex].BGNYDMJ / 10000, 4);

                cells[index, 26].Value = Math.Round(pdAtts[rowIndex].BGGDMJ / 10000, 4);
                cells[index, 27].Value = Math.Round(pdAtts[rowIndex].BGWLYDMJ / 10000, 4);
                cells[index, 28].Value = Math.Round(pdAtts[rowIndex].BGCZMJ / 10000, 4);
                cells[index, 29].Value = Math.Round(pdAtts[rowIndex].BGCKMJ / 10000, 4);
                cells[index, 30].Value = Math.Round(pdAtts[rowIndex].BGQTJSMJ / 10000, 4);

                cells[index, 31].Value = pdFeature["BZ"]?.ToString();
                cells[index, 32].Value = pdAtts[rowIndex].GHYDFLMC;

                // 写入固定值
                cells[index, 0].Value = index - 6;
                cells[index, 1].Value = "重庆市";

                string year = cells[index, 10].StringValue;
                cells[index, 11].Value = year[..4];

                index++;
                rowIndex++;
            }

            cells.DeleteRow(6);

            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        private void CreateExcel10(string pd, string gd)
        {
            // 获取原始图层和标识图层的要素类
            FeatureClass pdFeatureClass = pd.TargetFeatureClass();
            FeatureClass gdFeatureClass = gd.TargetFeatureClass();

            // 获取目标图层和源图层的要素游标
            using RowCursor pdCursor = pdFeatureClass.Search();

            // 遍历源图层的要素，标记
            while (pdCursor.MoveNext())
            {
                using Feature pdFeature = (Feature)pdCursor.Current;

                // 获取源要素的几何
                Polygon pdGeometry = pdFeature.GetShape() as Polygon;

                double pdPolygonArea = pdGeometry.Area;

                double totalArea = 0;

                List<string> pws = new List<string>();

                // 创建空间查询过滤器，以获取与源要素有重叠的目标要素
                SpatialQueryFilter spatialFilter = new SpatialQueryFilter
                {
                    FilterGeometry = pdGeometry,
                    SpatialRelationship = SpatialRelationship.Intersects
                };

                // 在目标图层中查询与源要素重叠的要素
                using RowCursor gdCursor = gdFeatureClass.Search(spatialFilter);
                while (gdCursor.MoveNext())
                {
                    using Feature gdFeature = (Feature)gdCursor.Current;
                    // 获取目标要素的几何
                    Polygon gdGeometry = gdFeature.GetShape() as Polygon;

                    // 计算源要素与目标要素的重叠面积
                    Geometry intersection = GeometryEngine.Instance.Intersection(pdGeometry, gdGeometry);

                    // 如果重叠面积大于当前最大重叠面积，则更新最大重叠面积和目标要素
                    if (intersection is not null)
                    {
                        double area = (intersection as Polygon).Area;
                        totalArea += area;
                        pws.Add(gdFeature["批复文"]?.ToString());
                    }
                }

                if (totalArea / pdPolygonArea > 0.95)
                {
                    string pwStr = "";
                    foreach (string pw in pws)
                    {
                        pwStr += $"{pw},";
                    }
                    if (true)
                    {
                        pdFeature["SJGYPW"] = pwStr[..^1];
                    }
                    pdFeature["QLLX"] = "060100";
                }

                pdFeature.Store();
            }

        }

        private void CreateExcel11(string excelPath1, string pd, string gd)
        {
            // 获取原始图层和标识图层的要素类
            FeatureClass pdFeatureClass = pd.TargetFeatureClass();

            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excelPath1);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath1);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 获取cells
            Cells cells = sheet.Cells;


            int index = 1;
            using RowCursor pdCursor = pdFeatureClass.Search();
            // 遍历源图层的要素，写入表格
            while (pdCursor.MoveNext())
            {
                using Feature pdFeature = (Feature)pdCursor.Current;

                List<string> fds = new List<string>()
                {
                    "FID","ID","SEC","SZXZQDM","SZXZQMC","XZQDM","XZQMC","XMMC","XMYDLX",
                    "XFPFWH","PFWH","PFSJ","PZNR","PZZMJ","XZJSYDMJ","NZYMJ","GDMJ",
                    "WLYDMJ","ZSTDMJ","ZZLX","JHLX","TDZSJZ","XTGYPW","SJGYPW","QLLX",
                    "PWSJLY","SM","BJBGYKNMJ","ID_1","SEC_1","DYPWSPCJ","BNJSLYQX",
                    "ZYYNMJ","ZYSTHXMJ","ZYCZKFBJMJ","BGZMJ","BGNYDMJ","BGGDMJ",
                    "BGWLYDMJ","BGCZMJ","BGCKMJ","BGQTJSMJ","NCJTJSYDLX","FJSYDLX",
                    "BGGYTDMJ","BGJTTDMJ","BDCGYTDMJ","BDCJTTDMJ","JTQX","BZ","Shape_Leng","Shape_Area"
                };

                for (int i = 0; i < fds.Count; i++)
                {
                    string str = pdFeature[fds[i]]?.ToString();
                    if (str is not null)
                    {
                        cells[index, i].Value = str;
                    }
                }

                index++;
            }

            // 保存
            wb.Save(excelFile);
            wb.Dispose();

        }

        private List<string> CheckData()
        {
            List<string> result = new List<string>();

            return result;
        }

        private void combox_pd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_pd);
        }

        private void combox_gd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_gd);
        }

        private void combox_kfbj_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_kfbj);
        }

        private void combox_sthx_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_sthx);
        }

        private void combox_jbnt_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_jbnt);
        }

        private void combox_bgcc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_bgcc);
        }

        private void combox_czc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_czc);
        }

        private void combox_fq_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fq);
        }
    }


    class pdAtt
    {
        public string OID { get; set; }
        public string PFSJ { get; set; }

        public double ZYYNMJ { get; set; } = 0;
        public double ZYSTHXMJ { get; set; } = 0;
        public double ZYCZKFBJMJ { get; set; } = 0;

        public double BGZMJ { get; set; } = 0;
        public double BGNYDMJ { get; set; } = 0;
        public double BGGDMJ { get; set; } = 0;
        public double BGWLYDMJ { get; set; } = 0;

        public double BGCZMJ { get; set; } = 0;
        public double BGCKMJ { get; set; } = 0;
        public double BGQTJSMJ { get; set; } = 0;

        public string GHYDFLMC { get; set; } = "";
    }
}
