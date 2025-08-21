using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
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
using Polygon = ArcGIS.Core.Geometry.Polygon;
using SpatialReference = ArcGIS.Core.Geometry.SpatialReference;

namespace CCTool.Scripts.CusTool2
{
    /// <summary>
    /// Interaction logic for ExportBoundary.xaml
    /// </summary>
    public partial class ExportBoundary : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public ExportBoundary()
        {
            InitializeComponent();

            textCADPath.Text = BaseTool.ReadValueFromReg("ExportBoundary", "in_cad");
            textExcelPath.Text = BaseTool.ReadValueFromReg("ExportBoundary", "out_folder");

            Init(combox_sr);
            Init(combox_sr2);

            combox_sr2.Visibility = Visibility.Hidden;
            tt.Visibility = Visibility.Hidden;  
        }

        // 初始化
        private void Init(ComboBox combox_sr)
        {
            // combox_sr框中添加几种预制坐标系
            combox_sr.Items.Add("(4490)GCS_China_Geodetic_Coordinate_System_2000");
            combox_sr.Items.Add("(4513)CGCS2000_3_Degree_GK_Zone_25");
            combox_sr.Items.Add("(4514)CGCS2000_3_Degree_GK_Zone_26");
            combox_sr.Items.Add("(4515)CGCS2000_3_Degree_GK_Zone_27");
            combox_sr.Items.Add("(4516)CGCS2000_3_Degree_GK_Zone_28");
            combox_sr.Items.Add("(4517)CGCS2000_3_Degree_GK_Zone_29");
            combox_sr.Items.Add("(4518)CGCS2000_3_Degree_GK_Zone_30");
            combox_sr.Items.Add("(4519)CGCS2000_3_Degree_GK_Zone_31");
            combox_sr.Items.Add("(4520)CGCS2000_3_Degree_GK_Zone_32");
            combox_sr.Items.Add("(4521)CGCS2000_3_Degree_GK_Zone_33");
            combox_sr.Items.Add("(4522)CGCS2000_3_Degree_GK_Zone_34");
            combox_sr.Items.Add("(4523)CGCS2000_3_Degree_GK_Zone_35");
            combox_sr.Items.Add("(4524)CGCS2000_3_Degree_GK_Zone_36");
            combox_sr.Items.Add("(4525)CGCS2000_3_Degree_GK_Zone_37");
            combox_sr.Items.Add("(4526)CGCS2000_3_Degree_GK_Zone_38");
            combox_sr.Items.Add("(4527)CGCS2000_3_Degree_GK_Zone_39");
            combox_sr.Items.Add("(4528)CGCS2000_3_Degree_GK_Zone_40");
            combox_sr.Items.Add("(4529)CGCS2000_3_Degree_GK_Zone_41");
            combox_sr.Items.Add("(4530)CGCS2000_3_Degree_GK_Zone_42");
            combox_sr.Items.Add("(4531)CGCS2000_3_Degree_GK_Zone_43");
            combox_sr.Items.Add("(4532)CGCS2000_3_Degree_GK_Zone_44");
            combox_sr.Items.Add("(4533)CGCS2000_3_Degree_GK_Zone_45");

            combox_sr.Items.Add("(4534)CGCS2000_3_Degree_GK_CM_75E");
            combox_sr.Items.Add("(4535)CGCS2000_3_Degree_GK_CM_78E");
            combox_sr.Items.Add("(4536)CGCS2000_3_Degree_GK_CM_81E");
            combox_sr.Items.Add("(4537)CGCS2000_3_Degree_GK_CM_84E");
            combox_sr.Items.Add("(4538)CGCS2000_3_Degree_GK_CM_87E");
            combox_sr.Items.Add("(4539)CGCS2000_3_Degree_GK_CM_90E");
            combox_sr.Items.Add("(4540)CGCS2000_3_Degree_GK_CM_93E");
            combox_sr.Items.Add("(4541)CGCS2000_3_Degree_GK_CM_96E");
            combox_sr.Items.Add("(4542)CGCS2000_3_Degree_GK_CM_99E");
            combox_sr.Items.Add("(4543)CGCS2000_3_Degree_GK_CM_102E");
            combox_sr.Items.Add("(4544)CGCS2000_3_Degree_GK_CM_105E");
            combox_sr.Items.Add("(4545)CGCS2000_3_Degree_GK_CM_108E");
            combox_sr.Items.Add("(4546)CGCS2000_3_Degree_GK_CM_111E");
            combox_sr.Items.Add("(4547)CGCS2000_3_Degree_GK_CM_114E");
            combox_sr.Items.Add("(4548)CGCS2000_3_Degree_GK_CM_117E");
            combox_sr.Items.Add("(4549)CGCS2000_3_Degree_GK_CM_120E");
            combox_sr.Items.Add("(4550)CGCS2000_3_Degree_GK_CM_123E");
            combox_sr.Items.Add("(4551)CGCS2000_3_Degree_GK_CM_126E");
            combox_sr.Items.Add("(4552)CGCS2000_3_Degree_GK_CM_129E");
            combox_sr.Items.Add("(4553)CGCS2000_3_Degree_GK_CM_132E");
            combox_sr.Items.Add("(4554)CGCS2000_3_Degree_GK_CM_135E");

        }


        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "CAD导出界址点Excel";


        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.OpenDialogFolder();
        }

        private void openCADButton_Click(object sender, RoutedEventArgs e)
        {
            textCADPath.Text = UITool.OpenDialogCAD();
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string in_cad = textCADPath.Text;
                string out_folder = textExcelPath.Text;
                string def_gdb = Project.Current.DefaultGeodatabasePath;
                string sr = combox_sr.Text;
                string sr2 = combox_sr2.Text;

                // 判断参数是否选择完全
                if (in_cad == "" || out_folder == "" || combox_sr.Text =="")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                if (combox_sr2.Visibility== Visibility.Visible)
                {
                    if (combox_sr2.Text == "")
                    {
                        MessageBox.Show("有必选参数为空！！！");
                        return;
                    }
                }

                // 保存设置
                BaseTool.WriteValueToReg("ExportBoundary", "in_cad", in_cad);
                BaseTool.WriteValueToReg("ExportBoundary", "out_folder", out_folder);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart($"CAD数据预处理");
                    // 创建目标文件夹
                    bool isHaveTarget = Directory.Exists(out_folder);
                    if (isHaveTarget)
                    {
                        Directory.Delete(out_folder, true);
                    }
                    Directory.CreateDirectory(out_folder);
                    Directory.CreateDirectory($@"{out_folder}\本体范围");
                    Directory.CreateDirectory($@"{out_folder}\本体边界坐标");

                    string fc = $@"{def_gdb}\地籍面";
                    string fc2 = $@"{def_gdb}\地籍面2";

                    int srID = int.Parse(sr[1..5]);
                    SpatialReference spatialReference = SpatialReferenceBuilder.CreateSpatialReference(srID);

                    // 给CAD定义坐标系
                    Arcpy.DefineProjection(in_cad, spatialReference);

                    // 先提取面积备用
                    List<string> areaList = new List<string>();
                    // 如果本身坐标系是地理坐标系
                    if (combox_sr2.Visibility == Visibility.Visible)
                    {
                        Arcpy.SpatialJoin($@"{in_cad}\Polygon", $@"{in_cad}\Annotation", fc2);
                        
                        int srID2 = int.Parse(sr2[1..5]);
                        SpatialReference spatialReference2 = SpatialReferenceBuilder.CreateSpatialReference(srID2);
                        Arcpy.Project(fc2, fc, spatialReference2, spatialReference);
                        areaList = fc.GetAllFieldValues("Shape_Area");
                    }
                    else
                    {
                        Arcpy.SpatialJoin($@"{in_cad}\Polygon", $@"{in_cad}\Annotation", fc);

                        SpatialReference spatialReference2 = SpatialReferenceBuilder.CreateSpatialReference(4490);
                        Arcpy.Project(fc, fc2, spatialReference2, spatialReference);
                        areaList = fc.GetAllFieldValues("Shape_Area");
                    }

                    // 遍历面要素类中的所有要素
                    FeatureClass featureClass = fc2.TargetFeatureClass();
                    RowCursor cursor = featureClass.Search();
                    int index = 0;
                    while (cursor.MoveNext())
                    {
                        using Feature feature = cursor.Current as Feature;
                        string feature_name = feature["TxtMemo"]?.ToString();
                        string elevation = double.Parse(feature["Elevation"]?.ToString()).RoundWithFill(3);

                        pw.AddMessageMiddle(10, $"处理要素_{feature_name}");

                        // 获取要素的几何
                        Polygon polygon = feature.GetShape() as Polygon;

                        // 复制Excel文件
                        string oriPath01 = $@"CCTool.Data.Excel.XXX本体范围导入模板.xls";
                        string oriPath02 = $@"CCTool.Data.Excel.XXX本体边界坐标导入模板.xls";
                        string targetExcel01 = $@"{out_folder}\本体范围\{feature_name}本体范围.xls";
                        string targetExcel02 = $@"{out_folder}\本体边界坐标\{feature_name}本体边界坐标.xls";
                        DirTool.CopyResourceFile(oriPath01, targetExcel01);
                        DirTool.CopyResourceFile(oriPath02, targetExcel02);

                        double area =double.Parse( areaList[index]);
                        MapPoint centerPoint = GeometryEngine.Instance.Centroid(polygon);   // 中心点

                        if (area > 10)      // 面积10平方米以上导出界址点
                        {
                            // 界址点重排
                            List<List<MapPoint>> mapPoints = polygon.ReshotMapPoint(false);
                            // 写入本体范围
                            WriteExcelBound(targetExcel01, mapPoints);
                            // 写入边界坐标
                            WriteExcelPoint(targetExcel02, mapPoints, centerPoint, elevation);

                        }
                        else    // 面积小于10平方米以上导出中心点
                        {
                            // 写入本体范围
                            WriteExcelBound(targetExcel01, centerPoint);
                            // 写入边界坐标
                            WriteExcelPoint(targetExcel02, centerPoint, elevation);
                        }
                        index++;
                    }

                    // 删除中间数据
                    Arcpy.Delect(fc);
                    Arcpy.Delect(fc2);

                });

                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }


        // 写入Excel本体范围
        public void WriteExcelBound(string excelPath, List<List<MapPoint>> mapPoints)
        {
            // 写入Excel
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 逐行处理
            int index = 1;
            foreach (List<MapPoint> mapPoint in mapPoints)
            {
                // 复制行
                if (index > 1)
                {
                    sheet.Cells.CopyRow(sheet.Cells, 1, index);
                }

                foreach (MapPoint pt in mapPoint)
                {
                    // 坐标点转换
                    string xx = ChangeDegree(pt.X);
                    string yy = ChangeDegree(pt.Y);
                    // 写入
                    sheet.Cells[index, 0].Value = (index - 1).ToString();
                    sheet.Cells[index, 1].Value = $@"测试点位{index}";
                    sheet.Cells[index, 2].Value = yy;
                    sheet.Cells[index, 3].Value = xx;

                    index++;
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // 写入Excel本体范围
        public void WriteExcelBound(string excelPath, MapPoint mapPoint)
        {
            // 写入Excel
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 逐行处理
            // 坐标点转换
            string xx = ChangeDegree(mapPoint.X);
            string yy = ChangeDegree(mapPoint.Y);

            // 写入
            for (int i = 1; i < 4; i++)
            {
                if (i > 1)
                {
                    // 复制行
                    sheet.Cells.CopyRow(sheet.Cells, 1, i);
                }

                sheet.Cells[i, 0].Value = (i - 1).ToString();
                sheet.Cells[i, 1].Value = $@"测试点位{i}";
                sheet.Cells[i, 2].Value = yy;
                sheet.Cells[i, 3].Value = xx;
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }


        // 写入Excel边界坐标
        public void WriteExcelPoint(string excelPath, List<List<MapPoint>> mapPoints, MapPoint centerPoint, string elevation)
        {
            // 写入Excel
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 逐行处理
            int index = 1;
            foreach (List<MapPoint> mapPoint in mapPoints)
            {
                foreach (MapPoint pt in mapPoint)
                {
                    // 复制行
                    if (index > 1)
                    {
                        sheet.Cells.CopyRow(sheet.Cells, 1, index);
                    }

                    // 坐标点转换
                    string xx = ChangeDegree(pt.X);
                    string yy = ChangeDegree(pt.Y);
                    // 高程值随机
                    Random random = new Random();
                    int randomNumber = random.Next(-1000, 1001);

                    string ele = (double.Parse(elevation) + randomNumber/1000.0).RoundWithFill(3);

                    // 写入
                    sheet.Cells[index, 0].Value = "边界点";
                    sheet.Cells[index, 1].Value = yy;
                    sheet.Cells[index, 2].Value = xx;
                    sheet.Cells[index, 3].Value = ele;
                    sheet.Cells[index, 4].Value = CalculateAngle(centerPoint, pt);    // 判断方位
                    sheet.Cells[index, 5].Value = "室外采集";

                    index++;
                }
            }
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }


        // 写入Excel边界坐标
        public void WriteExcelPoint(string excelPath, MapPoint mapPoint, string elevation)
        {
            // 写入Excel
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];
            // 逐行处理
            // 坐标点转换
            string xx = ChangeDegree(mapPoint.X);
            string yy = ChangeDegree(mapPoint.Y);
            // 写入
            sheet.Cells[1, 0].Value = "中心点";
            sheet.Cells[1, 1].Value = yy;
            sheet.Cells[1, 2].Value = xx;
            sheet.Cells[1, 3].Value = elevation;
            sheet.Cells[1, 4].Value = "中心点";
            sheet.Cells[1, 5].Value = "室外采集";
            // 保存
            wb.Save(excelFile);
            wb.Dispose();
        }

        // 坐标转换
        public string ChangeDegree(double value)
        {
            string result = "";

            // 计算度分秒的值
            int degree = (int)(value / 1);
            int minutes = (int)(value % 1 * 60 / 1);
            double seconds = (value % 1 * 60 - minutes) * 60;
            // 合并为字符串
            result = degree.ToString() + "°" + minutes.ToString() + "′" + seconds.ToString("0.0000") + "″";

            return result;
        }


        // 判断方位
        public string CalculateAngle(MapPoint centerPoint, MapPoint initPoint)
        {
            string result = "";

            if (initPoint.X > centerPoint.X && initPoint.Y > centerPoint.Y)
            {
                result = "东北角";
            }
            else if (initPoint.X > centerPoint.X && initPoint.Y < centerPoint.Y)
            {
                result = "东南角";
            }
            else if (initPoint.X < centerPoint.X && initPoint.Y > centerPoint.Y)
            {
                result = "西北角";
            }
            else if (initPoint.X < centerPoint.X && initPoint.Y < centerPoint.Y)
            {
                result = "西南角";
            }

            return result;
        }

        private void combox_sr_Closed(object sender, EventArgs e)
        {
            if (combox_sr.Text == "(4490)GCS_China_Geodetic_Coordinate_System_2000")
            {
                combox_sr2.Visibility = Visibility.Visible;
                tt.Visibility = Visibility.Visible;
            }
            else
            {
                combox_sr2.Visibility = Visibility.Hidden;
                tt.Visibility = Visibility.Hidden;
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/143854222";
            UITool.Link2Web(url);
        }
    }
}
