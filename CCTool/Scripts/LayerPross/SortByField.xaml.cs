using ActiproSoftware.Windows.Shapes;
using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using MathNet.Numerics.LinearAlgebra.Factorization;
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
using Row = ArcGIS.Core.Data.Row;

namespace CCTool.Scripts.LayerPross
{
    /// <summary>
    /// Interaction logic for SortByField.xaml
    /// </summary>
    public partial class SortByField : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "SortByField";

        public SortByField()
        {
            InitializeComponent();

            comBox_model.Items.Add("左上-->右下");
            comBox_model.Items.Add("右下-->左上");
            comBox_model.SelectedIndex = 0;

            // 初始化其它参数选项
            txt_length.Text = BaseTool.ReadValueFromReg(toolSet, "txt_length");

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "分区排序";

        private async void btn_go_click(object sender, RoutedEventArgs e)
        {
            // 获取指标
            string sortField = combox_field.ComboxText();
            string resultField = combox_resultField.ComboxText();
            string model = comBox_model.Text;
            int start =int.Parse(txt_bh.Text);

            _ = int.TryParse( txt_length.Text, out int fdLength);

            // 判断fdLength是否正确
            if (fdLength == 0)
            {
                MessageBox.Show("输入的编号长度有误！！！");
                return;
            }

            // 获取默认数据库
            var gdb = Project.Current.DefaultGeodatabasePath;

            // 判断参数是否选择完全
            if (resultField == "")
            {
                MessageBox.Show("有必选参数为空！！！");
                return;
            }

            // 保存参数
            BaseTool.WriteValueToReg(toolSet, "txt_length", txt_length.Text);

            // 打开进度框
            ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
            pw.AddMessageTitle(tool_name);

            Close();

            await QueuedTask.Run(() =>
            {

                pw.AddMessageStart("排序");
                // 获取图层
                FeatureLayer ly = MapView.Active.GetSelectedLayers().FirstOrDefault() as FeatureLayer;

                // 获取图层要素的路径
                string fc_path = ly.Name.TargetLayerPath();

                // 如果设置的长度超过字段原本长度，就按原长度
                Field field = GisTool.GetFieldFromString(fc_path, resultField);
                if (field.Length<fdLength)
                {
                    fdLength = field.Length;
                }

                // 获取符号系统
                CIMRenderer cr = ly.GetRenderer();
                // 模式转换
                string md = model switch
                {
                    "左上-->右下"=> "UL",
                    "右下-->左上" => "LR",
                    _=>null,
                };
                // 排序
                Arcpy.Sort(ly, gdb + @"\sort_fc", "Shape ASCENDING", md);
                // 更新要素
                Arcpy.CopyFeatures(gdb + @"\sort_fc", fc_path, true);
                // 应用符号系统
                ly.SetRenderer(cr);

                pw.AddMessageMiddle(30, "分区赋值");
                // 设置一个字典保存个数
                Dictionary<string, int> dic = new Dictionary<string, int>();
                List<string> values = new List<string>() { "默认值"};
                // 如果选择了分区字段
                if (sortField!="")
                {
                    // 获取所有字段值
                    values = ly.GetFieldValues(sortField);
                }

                foreach (string value in values)
                {
                    dic.Add(value, start);
                }
                // 获取ID字段
                string oidField = fc_path.TargetIDFieldName();
                // 分区赋值
                RowCursor cursor = ly.Search();
                while (cursor.MoveNext())
                {
                    using Row row = cursor.Current;

                    if (sortField != "")
                    {
                        // 分区名称
                        string area = row[sortField].ToString();
                        // 如果在分区列表中，就更新数量
                        if (dic.ContainsKey(area))
                        {
                            row[resultField] = dic[area].ToString().PadLeft(fdLength, '0') ;
                            dic[area] += 1;
                        }
                    }
                    else    // 如果没有选择分区字段
                    {
                        row[resultField] = dic["默认值"].ToString().PadLeft(fdLength, '0'); ;
                        dic["默认值"] += 1;
                    }
                    
                    // 保存
                    row.Store();
                }
            });

            pw.AddMessageEnd();

        }


        private void combox_field_DropOpen(object sender, EventArgs e)
        {
            FeatureLayer ly = MapView.Active.GetSelectedLayers().FirstOrDefault() as FeatureLayer;
            UITool.AddTextFieldsToComboxPlus(ly, combox_field);
        }

        private void btn_help_click(object sender, RoutedEventArgs e)
        {
            string url = "";
            UITool.Link2Web(url);
        }

        private void combox_resultField_DropOpen(object sender, EventArgs e)
        {
            FeatureLayer ly = MapView.Active.GetSelectedLayers().FirstOrDefault() as FeatureLayer;
            UITool.AddIntAndTextFieldsToComboxPlus(ly, combox_resultField);
        }
    }
}
