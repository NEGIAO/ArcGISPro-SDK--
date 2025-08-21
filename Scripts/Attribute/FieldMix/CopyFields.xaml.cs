using ArcGIS.Core.Data;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing.Attributes;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
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
using Table = ArcGIS.Core.Data.Table;

namespace CCTool.Scripts.UI.ProWindow
{
    /// <summary>
    /// Interaction logic for CopyFields.xaml
    /// </summary>

    // 创建一个字段类
    public class FieldDef
    {
        public string fldName { get; set; }
        public string fldAlias { get; set; }
        public string fldType { get; set; }
        public int fldLength { get; set; }
    }


    public partial class CopyFields : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public CopyFields()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "复制字段";

        private void combox_fc_before_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayerAndTableToComboxPlus(combox_fc_before);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                pw.AddMessageStart("获取相关参数");

                // 参数获取
                string fc_before = combox_fc_before.ComboxText();
                var fc_after = listbox_targetFeature.Items;
                var fileds = listbox_field.Items;

                // 判断参数是否选择完全
                if (fc_before == "" || listbox_targetFeature.Items.Count == 0 || listbox_field.Items.Count == 0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }
                Close();

                // 获取参数listbox
                List<string> fieldNames = UITool.GetCheckboxStringFromListBox(listbox_field);
                List<string> targetFeatureClasses = UITool.GetCheckboxStringFromListBox(listbox_targetFeature);

                await QueuedTask.Run(() =>
                {
                    List<FieldDef> fieldDefs = new List<FieldDef>();

                    pw.AddMessageMiddle(20, @"获取字段属性");
                    // 获取字段属性
                    foreach (string fieldName in fieldNames)
                    {
                        FieldDef fd = new FieldDef();
                        Table table = fc_before.TargetTable();
                        var inspector = new Inspector();
                        inspector.LoadSchema(table);
                        // 获取属性
                        foreach (var att in inspector)
                        {
                            // 如果符合字段名
                            if (att.FieldName == fieldName)
                            {
                                fd.fldName = att.FieldName;
                                fd.fldAlias = att.FieldAlias;
                                fd.fldType = att.FieldType.ToString();
                                fd.fldLength = att.Length;
                            }
                        }
                        // 加入字段集合
                        fieldDefs.Add(fd);
                    }

                    // 复制字段
                    foreach (string targetFeatureClass in targetFeatureClasses)
                    {
                        foreach (var fd in fieldDefs)
                        {
                            if (!GisTool.IsHaveFieldInTarget(targetFeatureClass, fd.fldName))
                            {
                                pw.AddMessageMiddle(10, $"【{targetFeatureClass}】__ 复制字段：{fd.fldName}");
                                Arcpy.AddField(targetFeatureClass, fd.fldName, fd.fldType, fd.fldAlias, fd.fldLength);
                            }
                            else
                            {
                                pw.AddMessageMiddle(10, $"【{targetFeatureClass}】已经存在字段：{fd.fldName}", Brushes.IndianRed);
                            }
                            
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

        private void combox_fc_before_DropClose(object sender, EventArgs e)
        {
            try
            {
                string fcPath = combox_fc_before.ComboxText();

                // 生成字段列表
                if (fcPath != "")
                {
                    UITool.AddFieldsToListBox(listbox_field,fcPath);
                }

                // 将剩余要素图层放在目标图层中
                UITool.AddFeatureLayersAndTablesToListbox(listbox_targetFeature, new List<string>() { fcPath });

            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        private void btn_select_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_field);
        }

        private void btn_unSelect_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_field);
        }

        private void btn_select_fc_Click(object sender, RoutedEventArgs e)
        {
            UITool.SelectListboxItems(listbox_targetFeature);
        }

        private void btn_unSelect_fc_Click(object sender, RoutedEventArgs e)
        {
            UITool.UnSelectListboxlItems(listbox_targetFeature);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/135742751?spm=1001.2014.3001.5501";
            UITool.Link2Web(url);
        }

        private void combox_sd(object sender, EventArgs e)
        {

        }
    }
}
