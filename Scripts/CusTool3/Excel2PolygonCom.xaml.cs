using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
using CCTool.Scripts.ToolManagers.Managers;
using CCTool.Scripts.ToolManagers.Windows;
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

namespace CCTool.Scripts.CusTool3
{
    /// <summary>
    /// Interaction logic for Excel2PolygonCom.xaml
    /// </summary>
    public partial class Excel2PolygonCom : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "Excel2PolygonCom";


        public Excel2PolygonCom()
        {
            InitializeComponent();

            EventCenter.AddListener(EventDefine.UpdataCod, UpdataCod);

            // 初始化
            textFolderPath.Text = BaseTool.ReadValueFromReg(toolSet, "folderPath");
            textFeatureClassPath.Text = BaseTool.ReadValueFromReg(toolSet, "fcPath");
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
        string tool_name = "Excel点集转面要素(批量)";


        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取指标
                string folderPath = textFolderPath.Text;
                string fcPath = textFeatureClassPath.Text;
                string srName = textCod.Text;

                var cb_txts = listbox_txt.Items;

                // 参数写入本地
                BaseTool.WriteValueToReg(toolSet, "folderPath", folderPath);
                BaseTool.WriteValueToReg(toolSet, "fcPath", fcPath);
                BaseTool.WriteValueToReg(toolSet, "srName", srName);

                // 判断参数是否选择完全
                if (folderPath == "" || fcPath == "" || cb_txts.Count == 0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 获取目标数据库和点要素名
                string gdbPath = fcPath.TargetWorkSpace();
                string fcName = fcPath.TargetFcName();

                // 获取所有选中的excel
                List<string> list_excelPath = new List<string>();
                foreach (CheckBox shp in cb_txts)
                {
                    if (shp.IsChecked == true)
                    {
                        list_excelPath.Add(folderPath + shp.Content);
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
                    // 检查数据
                    List<string> errs = CheckData(fcPath);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    // 参数获取
                    string gdb_def = Project.Current.DefaultGeodatabasePath;
                    pw.AddMessageMiddle(10, "创建一个空要素");

                    // 创建一个空要素
                    Arcpy.CreateFeatureclass(gdbPath, fcName, "POLYGON", srName);

                    // 新建字段
                    GisTool.AddField(fcPath, "项目名称");

                    pw.AddMessageMiddle(10, "获取所有Excel文件");

                    // 打开数据库
                    using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdb_def))))
                    {
                        // 创建要素并添加到要素类中
                        using FeatureClass featureClass = gdb.OpenDataset<FeatureClass>(fcName);
                        // 解析excel文件内容，创建面要素
                        foreach (string path in list_excelPath)
                        {
                            // 获取项目名称，并处理掉一些干扰文字
                            string featureName = path[(path.LastIndexOf(@"\") + 1)..].Replace(".xls","").Replace(".xlsx", "").Replace("_转自XLSX", "");
                            pw.AddMessageMiddle(1, "       处理：" + featureName, Brushes.Gray);

                            // 收集点集
                            var vertices = new List<Coordinate2D>();
                            // 获取工作薄、工作表
                            string excelFile =ExcelTool.GetPath(path);
                            int sheetIndex = ExcelTool.GetSheetIndex(path);
                            // 打开工作薄
                            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                            // 打开工作表
                            Worksheet sheet = wb.Worksheets[sheetIndex];
                            Cells cells= sheet.Cells;

                            // 逐行处理
                            for (int i = 1; i <= sheet.Cells.MaxDataRow; i++)
                            {
                                //  获取目标cell
                                Cell xCell = sheet.Cells[i, 1];
                                Cell yCell = sheet.Cells[i, 2];
                                // 属性映射
                                if (xCell is not null && yCell is not null)
                                {
                                    vertices.Add(new Coordinate2D(xCell.DoubleValue, yCell.DoubleValue));
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
                                rowBuffer["项目名称"] = featureName;

                                PolygonBuilderEx pb = new PolygonBuilderEx(vertices);

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
                    // 保存编辑
                    Project.Current.SaveEditsAsync();

                    // 修复几何
                    Arcpy.RepairGeometry(fcPath);

                    // 将要素类添加到当前地图
                    MapCtlTool.AddLayerToMap(fcPath);

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
            textFeatureClassPath.Text = UITool.SaveDialogFeatureClass();
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
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
            if (folder != "")
            {
                // 获取所有excel文件
                List<string> excelFiles = new List<string>() { ".xls" , ".xlsx" };
                var files = DirTool.GetAllFilesFromList(folder, excelFiles);
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


        private List<string> CheckData(string featurePath)
        {
            List<string> result = new List<string>();

            // 判断输出路径是否为gdb
            string gdbResult = CheckTool.CheckGDBPath(featurePath);
            if (gdbResult != "")
            {
                result.Add(gdbResult);
            }

            // 检查gdb要素是否是以数字开关
            string numbericResult = CheckTool.CheckGDBIsNumeric(featurePath);
            if (numbericResult != "https://blog.csdn.net/xcc34452366/article/details/145229311?spm=1001.2014.3001.5502")
            {
                result.Add(numbericResult);
            }
            return result;
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "";
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
    }
}
