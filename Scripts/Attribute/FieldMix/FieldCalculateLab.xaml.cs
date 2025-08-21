using ArcGIS.Desktop.Core;
using Aspose.Cells;
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

namespace CCTool.Scripts.Attribute.FieldMix
{
    /// <summary>
    /// Interaction logic for FieldCalculateLab.xaml
    /// </summary>
    public partial class FieldCalculateLab : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public FieldCalculateLab()
        {
            InitializeComponent();

            // 加载上次的数据
            CalculateBox.Text = BaseTool.ReadValueFromReg(toolSet, "CalculateBox");
            Block.Text = BaseTool.ReadValueFromReg(toolSet, "Block");
            //explain.Content = BaseTool.ReadValueFromReg(toolSet, "explain");
            explain.Text = BaseTool.ReadValueFromReg(toolSet, "explain");
            // 复制字段计算器表
            string excelPath = @$"{Project.Current.HomeFolderPath}\字段计算器公式.xlsx";
            DirTool.CopyResourceFile(@"CCTool.Data.Excel.字段计算器公式.xlsx", excelPath);
        }

        // 工具设置标签
        readonly string toolSet = "FieldCalculateLab";


        private void itemClick(object sender, RoutedEventArgs e)
        {
            // 获取Button
            Button button = sender as Button;

            if (button is null)
            {
                return;
            }

            //  名称
            string TName = button.Content.ToString();

            // 获取字段计算器
            List<CalAtt> calAtts = new List<CalAtt>();

            string excelPath = @$"{Project.Current.HomeFolderPath}\字段计算器公式.xlsx";

            // 获取工作薄、工作表
            string excelFile = ExcelTool.GetPath(excelPath);
            int sheetIndex = ExcelTool.GetSheetIndex(excelPath);
            // 打开工作薄
            Workbook wb = ExcelTool.OpenWorkbook(excelFile);
            // 打开工作表
            Worksheet sheet = wb.Worksheets[sheetIndex];

            Cells cells = sheet.Cells;

            // 逐行处理
            for (int i = 1; i <= cells.MaxDataRow; i++)
            {
                //  名称
                string tName = cells[i, 1].StringValue;
                //  表达式
                string CalculateBox = cells[i, 2].StringValue;
                //  代码块
                string Block = cells[i, 3].StringValue;
                //  说明
                string explain = cells[i, 4].StringValue;

                // 写入
                CalAtt calAtt = new CalAtt()
                {
                    Name = tName,
                    Expression= CalculateBox,
                    Block= Block,
                    Explain= explain
                };

                calAtts.Add(calAtt);

            }
            wb.Dispose();


            

            // 更新工具面板
            foreach (var calAtt in calAtts)
            {
                if (calAtt.Name == button.Name)
                {
                    tName.Content= button.Content;
                    CalculateBox.Text = calAtt.Expression;
                    Block.Text = calAtt.Block;
                    explain.Text = calAtt.Explain;
                }
            }

            // 保存数据
            BaseTool.WriteValueToReg(toolSet, "CalculateBox", CalculateBox.Text);
            BaseTool.WriteValueToReg(toolSet, "Block", Block.Text);
            BaseTool.WriteValueToReg(toolSet, "explain", explain.Text);

        }


        private void pw_Unload(object sender, RoutedEventArgs e)
        {
            // 删除字段计算器表
            string excelPath = @$"{Project.Current.HomeFolderPath}\字段计算器公式.xlsx";
            File.Delete(excelPath);
        }
    }



}

public class CalAtt
{
    public string Name { get; set; }
    public string Expression { get; set; }
    public string Block { get; set; }
    public string Explain { get; set; }
}

