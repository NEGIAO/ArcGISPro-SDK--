using ArcGIS.Core.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
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

namespace CCTool.Scripts.MixApp.StyleMix
{
    /// <summary>
    /// Interaction logic for CreateSimplePolygonStyle.xaml
    /// </summary>
    public partial class CreateSimplePolygonStyle : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "CreateSimplePolygonStyle";
        public CreateSimplePolygonStyle()
        {
            InitializeComponent();

            // 加载保存的设置
            textExcelPath.Text = BaseTool.ReadValueFromReg(toolSet, "excelPath");
            textName.Text = BaseTool.ReadValueFromReg(toolSet, "stylxName");

        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string stylxName = textName.Text;
                string excelPath = textExcelPath.Text;

                // 判断参数是否选择完全
                if (stylxName == "" || excelPath == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 设置保存在本地
                BaseTool.WriteValueToReg(toolSet, "excelPath", excelPath);
                BaseTool.WriteValueToReg(toolSet, "stylxName", stylxName);

                Close();
                await QueuedTask.Run(() =>
                {
                    // 创建样式
                    StylxTool.CreateStylx(stylxName);
                    // 获取样式
                    StyleProjectItem styleProjectItem = stylxName.TargetStyleProjectItem();

                    // 获取工作薄、工作表
                    string excelFile = ExcelTool.GetPath(excelPath);
                    int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
                    // 打开工作薄
                    Workbook wb = ExcelTool.OpenWorkbook(excelFile);
                    // 打开工作表
                    Worksheet sheet = wb.Worksheets[sheetIndex];

                    Cells cells = sheet.Cells;

                    // 逐行获取
                    for (int i = 1; i <= sheet.Cells.MaxDataRow; i++)
                    {
                        // 名称，标签，RGB
                        string name = cells[i, 0].StringValue;
                        string tag = cells[i, 1].StringValue;
                        string rgbValue = cells[i, 2].StringValue;
                        // 分解RGB
                        int r = rgbValue.Split(';')[0].ToInt();
                        int g = rgbValue.Split(';')[1].ToInt();
                        int b = rgbValue.Split(';')[2].ToInt();

                        // 判断一下
                        if (name is null || name == ""|| tag is null || tag == "")
                        {
                            continue;
                        }
                        // 轮廓宽度和颜色
                        string lineLen = cells[i, 3].StringValue;
                        double lineLength = lineLen is null ? 0 : lineLen.ToDouble();

                        string lineRGBValue = cells[i, 4].StringValue;

                        // 分解轮廓RGB
                        int r_line = lineRGBValue.Split(';')[0].ToInt();
                        int g_line = lineRGBValue.Split(';')[1].ToInt();
                        int b_line = lineRGBValue.Split(';')[2].ToInt();

                        // 创建RGB颜色
                        CIMColor rgbColor = ColorFactory.Instance.CreateRGBColor(r, g, b);

                        //// 创建CMYK颜色对象
                        //CIMCMYKColor cmykColor = new CIMCMYKColor()
                        //{
                        //    C = 20,    // 青色分量 (0-100)
                        //    M = 60,    // 洋红色分量 (0-100)
                        //    Y = 30,    // 黄色分量 (0-100)
                        //    K = 10,    // 黑色分量 (0-100)
                        //    Alpha = 100 // 不透明度 (0-100, 100=完全不透明)
                        //};


                        CIMColor lineRGBColor = ColorFactory.Instance.CreateRGBColor(r_line, g_line, b_line);

                        // 轮廓线
                        var outline = SymbolFactory.Instance.ConstructStroke(lineRGBColor,lineLength,SimpleLineStyle.Solid);

                        // 面样式
                        CIMPolygonSymbol polySymbol = SymbolFactory.Instance.ConstructPolygonSymbol(rgbColor, SimpleFillStyle.Solid,outline);

                        // 创建符号样式项
                        var symbolItem = new SymbolStyleItem()
                        {
                            Name = name,  //符号名称
                            Category = "", //分类目录
                            Tags = tag, //标签
                            Key = name,  //唯一标识（可选）
                            Symbol = polySymbol,
                        };
                        // 将符号添加到样式库
                        styleProjectItem.AddItem(symbolItem);

                    }

                    wb.Dispose();

                });

                MessageBox.Show($"创建简单面样式完成!");
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147196202";
            UITool.Link2Web(url);
        }

        private void downLoadButton_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://pan.baidu.com/s/1RrmXEHq5etIZ4Hv7WO8aZA?pwd=6vw8";
            UITool.Link2Web(url);
        }

        private void openExcelButton_Click(object sender, RoutedEventArgs e)
        {
            textExcelPath.Text = UITool.OpenDialogExcel();
        }
    }
}
