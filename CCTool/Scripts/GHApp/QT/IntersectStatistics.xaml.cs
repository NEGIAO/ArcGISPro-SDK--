using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
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

namespace CCTool.Scripts.GHApp.QT
{
    /// <summary>
    /// Interaction logic for IntersectStatistics.xaml
    /// </summary>
    public partial class IntersectStatistics : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public IntersectStatistics()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "相交占比分析";

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取默认数据库
                var def_gdb = Project.Current.DefaultGeodatabasePath;
                // 获取要素
                string origin = combox_origin.ComboxText();
                string identy = combox_identy.ComboxText();

                // 判断参数是否选择完全
                if (origin == "" || identy == "")
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
                    List<string> lines = new List<string>() { origin , identy };
                    // 检查数据
                    List<string> errs = CheckData(lines);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }


                    pw.AddMessageMiddle(20, "相交并标记");

                    // 标记一个ID字段
                    string oid = origin.TargetIDFieldName();
                    string field = "标记";
                    Arcpy.AddField(origin, field, "LONG");
                    Arcpy.CalculateField(origin, field, @$"!{oid}!");

                    List<string> in_fcs = new List<string>() { origin, identy };
                    string intersect = $@"{def_gdb}\intersect";
                    Arcpy.Intersect(in_fcs, intersect);

                    // 计算面积
                    GisTool.AddField(intersect, "面积标记", ArcGIS.Core.Data.FieldType.Double);
                    Arcpy.CalculateField(intersect, "面积标记", "!shape.area!");

                    pw.AddMessageMiddle(20, "汇总并连接字段");
                    // 先判断一下，如果有XJ_MJ字段，就先删除
                    if (GisTool.IsHaveFieldInTarget(origin, "XJ_MJ"))
                    {
                        Arcpy.DeleteField(origin, "XJ_MJ");
                    }
                    // 汇总并连接字段
                    string fd_area = "XJ_MJ";
                    string fd_per = "XJ_ZB";
                    string out_table = $@"{def_gdb}\out_table";
                    Arcpy.Statistics(intersect, out_table, "面积标记 SUM", field);
                    Arcpy.AlterField(out_table, "SUM_面积标记", fd_area, fd_area);
                    Arcpy.JoinField(origin, oid, out_table, field, new List<string>() { fd_area });

                    pw.AddMessageMiddle(20, "计算占比");
                    // 计算占比
                    Arcpy.AddField(origin, fd_per, "DOUBLE");
                    Arcpy.CalculateField(origin, fd_per, $"!{fd_area}!/!Shape.Area!");

                    // 获取中间字段
                    Arcpy.Delect(intersect);
                    Arcpy.Delect(out_table);
                    Arcpy.DeleteField(origin, field);

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
            string url = "https://blog.csdn.net/xcc34452366/article/details/140240696";
            UITool.Link2Web(url);
        }


        private void combox_origin_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_origin);
        }


        private void combox_identy_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_identy);
        }


        private List<string> CheckData(List<string> lines)
        {
            List<string> result = new List<string>();


            return result;
        }
    }
}
