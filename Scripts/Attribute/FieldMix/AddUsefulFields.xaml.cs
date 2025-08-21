using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.UI.ProMapTool;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    /// Interaction logic for AddUsefulFields.xaml
    /// </summary>
    public partial class AddUsefulFields : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public AddUsefulFields()
        {
            InitializeComponent();

            // 初始化combox
            combox_fieldGroup.Items.Add("通用");
            combox_fieldGroup.Items.Add("国土空间规划");
            combox_fieldGroup.Items.Add("三调");
            combox_fieldGroup.SelectedIndex = 0;

            // 将当前地图的要素图层和独立表加入到listbox
            UITool.AddFeatureLayersAndTablesToListbox(listbox_targetFeature);

            // 更新表格内容
            UpdataDG();
        }


        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "添加常用字段";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取图层数据
                List<string> tableList = UITool.GetCheckboxStringFromListBox(listbox_targetFeature);
                
                List<List<string>> fieldPairs = new List<List<string>>();
                // 获取字段属性
                for (int i = 0; i < dg.Items.Count; i++)
                {
                    // 获取选择框
                    CheckBox item_isCheck = (CheckBox)dg.GetCell(i, 0).Content;
                    bool isCheck = (bool)item_isCheck.IsChecked;
                    // 获取名称
                    TextBlock item_name = (TextBlock)dg.GetCell(i, 1).Content;
                    string name = item_name.Text;
                    // 获取别名
                    TextBlock item_aliasName = (TextBlock)dg.GetCell(i, 2).Content;
                    string aliasName = item_aliasName.Text;
                    // 获取类型
                    TextBlock item_type = (TextBlock)dg.GetCell(i, 3).Content;
                    string type = item_type.Text;
                    // 获取长度
                    TextBlock item_length = (TextBlock)dg.GetCell(i, 4).Content;
                    string length = item_length.Text;

                    // 如果是选中状态
                    if (isCheck)
                    {
                        fieldPairs.Add(new List<string>() { name, aliasName, type, length});
                    }
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                await QueuedTask.Run(() =>
                {
                    // 添加字段
                    pw.AddMessageStart("Start!!");
                    foreach (string table in tableList)
                    {
                        pw.AddMessageMiddle(10, $"处理图层：{table}");

                        foreach (var fieldPair in fieldPairs)
                        {
                            pw.AddMessageMiddle(1, $"      添加字段：{fieldPair[0]}", Brushes.Gray);

                            Arcpy.AddField(table, fieldPair[0], fieldPair[2], fieldPair[1], fieldPair[3].ToInt());
                        }
                    }
                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }


        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/148287909";
            UITool.Link2Web(url);
        }


        private void btn_select_fc_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_targetFeature);
        }

        private void btn_unSelect_fc_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_targetFeature);
        }


        // 定义一个空包
        List<FieldAtt> fieldAtt = new List<FieldAtt>();

        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            if (dg.Items.Count == 0)
            {
                MessageBox.Show("列表内没有字段！");
                return;
            }

            // 定义一个新包
            List<FieldAtt> fieldAtt2 = fieldAtt;

            foreach (var item in fieldAtt2)
            {
                item.IsCheck = true;
            }

            // 绑定
            dg.ItemsSource = fieldAtt2;
            // 刷新
            dg.Items.Refresh();
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            if (dg.Items.Count == 0)
            {
                MessageBox.Show("列表内没有字段！");
                return;
            }

            // 定义一个新包
            List<FieldAtt> fieldAtt2 = fieldAtt;

            foreach (var item in fieldAtt2)
            {
                item.IsCheck = false;

            }

            // 绑定
            dg.ItemsSource = fieldAtt2;
            // 刷新
            dg.Items.Refresh();
        }

        private void combox_fieldGroup_Closed(object sender, EventArgs e)
        {
            // 更新表格内容
            UpdataDG();
        }

        // 更新表格内容
        private void UpdataDG()
        {
            try
            {
                string fieldGroup = combox_fieldGroup.Text;

                List<List<string>> fields = new List<List<string>>();

                // 类别
                if (fieldGroup == "通用")
                {
                    fields.Add(["BSM", "标识码", "text", "18"]);
                    fields.Add(["YSDM", "要素代码", "text", "10"]);
                    fields.Add(["XZQDM", "行政区代码", "text", "12"]);
                    fields.Add(["XZQMC", "行政区名称", "text", "100"]);
                    fields.Add(["TBMJ", "图斑面积", "double", ""]);
                    fields.Add(["BZ", "备注", "text", "255"]);
                }
                else if (fieldGroup == "国土空间规划")
                {
                    fields.Add(["YDYHFLDM", "用地用海分类代码", "text", "10"]);
                    fields.Add(["YDYHFLMC", "用地用海分类名称", "text", "50"]);
                    fields.Add(["JQDLBM", "基期地类编码", "text", "10"]);
                    fields.Add(["JQDLMC", "基期地类名称", "text", "50"]);
                    fields.Add(["GHDLBM", "规划地类编码", "text", "10"]);
                    fields.Add(["GHDLMC", "规划地类名称", "text", "50"]);
                }
                else if (fieldGroup == "三调")
                {
                    fields.Add(["DLBM", "地类编码", "text", "5"]);
                    fields.Add(["DLMC", "地类名称", "text", "60"]);
                }

                // 定义一个空包
                List<FieldAtt> fieldAtt2 = new List<FieldAtt>();

                // 添加数据

                foreach (var field in fields)
                {
                    fieldAtt2.Add(new FieldAtt()
                    {
                        IsCheck = false,
                        FieldName = field[0],
                        AliasName = field[1],
                        FieldType = field[2],
                        FieldLength = field[3],
                    });
                }

                // 绑定
                dg.ItemsSource = fieldAtt2;

                // 赋值
                fieldAtt = fieldAtt2;

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        class FieldAtt
        {
            public bool IsCheck { get; set; }
            public string FieldName { get; set; }
            public string AliasName { get; set; }
            public string FieldType { get; set; }
            public string FieldLength { get; set; }
        }
    }
}
