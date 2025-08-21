using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using MathNet.Numerics.Distributions;
using Org.BouncyCastle.Crypto;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Geometry = ArcGIS.Core.Geometry.Geometry;

namespace CCTool.Scripts.CusTool2
{
    /// <summary>
    /// Interaction logic for Resort.xaml
    /// </summary>
    public partial class Resort : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public Resort()
        {
            InitializeComponent();

            string scale = BaseTool.ReadValueFromReg("Resort", "txtScale");
            txtScale.Text = scale == "" ? "10000" : scale;
            string width = BaseTool.ReadValueFromReg("Resort", "txtWidth");
            txtWidth.Text = width == "" ? "39" : width;
            string height = BaseTool.ReadValueFromReg("Resort", "txtHeight");
            txtHeight.Text = height == "" ? "20.72" : height;
            string count = BaseTool.ReadValueFromReg("Resort", "txtCount");
            txtCount.Text = count == "" ? "8" : count;
            string group = BaseTool.ReadValueFromReg("Resort", "txtGroup");
            txtGroup.Text = group == "" ? "1" : group;
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "小班打印编组";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void combox_xzq_DropDown(object sender, EventArgs e)
        {
            string fc_path = combox_fc.ComboxText();
            UITool.AddTextFieldsToComboxPlus(fc_path, combox_xzq);
        }

        private void combox_bh_DropDown(object sender, EventArgs e)
        {
            string fc_path = combox_fc.ComboxText();
            UITool.AddIntFieldsToComboxPlus(fc_path, combox_bh);
        }


        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                // 获取默认数据库
                var init_gdb = Project.Current.DefaultGeodatabasePath;
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string field_xzq = combox_xzq.ComboxText();
                string field_bh = combox_bh.ComboxText();

                int scale = int.Parse(txtScale.Text);
                double width = double.Parse(txtWidth.Text);
                double height = double.Parse(txtHeight.Text);
                int count = int.Parse(txtCount.Text);
                int gp = int.Parse(txtGroup.Text);

                // 计算真正的宽和高
                double xs = 0.9;     // 缩放系数
                double scWidth = width * scale / 100 * xs;
                double scHeight = height * scale / 100 * xs;

                // 判断参数是否选择完全
                if (fc_path == "" || field_xzq == "" || field_bh == "" || field_xzq == "" || scale == 0 || width == 0 || height == 0 || count == 0)
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入页面数据
                BaseTool.WriteValueToReg("Resort", "txtScale", scale);
                BaseTool.WriteValueToReg("Resort", "txtWidth", width);
                BaseTool.WriteValueToReg("Resort", "txtHeight", height);
                BaseTool.WriteValueToReg("Resort", "txtCount", count);
                BaseTool.WriteValueToReg("Resort", "txtGroup", gp);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    string defGDB = Project.Current.DefaultGeodatabasePath;
                    string fc = $@"{defGDB}\{fc_path}_小班排序";

                    pw.AddMessageStart( $"小班排序");

                    // 复制OID值作为连接字段
                    Arcpy.AddField(fc_path, "连接", "LONG");
                    string oidField = fc_path.TargetIDFieldName();
                    Arcpy.CalculateField(fc_path, "连接", $"!{oidField}!");
                    // 清除预设字段
                    Arcpy.DeleteField(fc_path, "分组");
                    Arcpy.DeleteField(fc_path, "比例尺");
                    Arcpy.DeleteField(fc_path, "备注");

                    // 排序
                    Arcpy.Sort(fc_path, fc, $"{field_xzq} ASCENDING;{field_bh} ASCENDING", "UR");

                    // 添加标记字段
                    GisTool.AddField(fc, "分组", FieldType.Integer);
                    GisTool.AddField(fc, "比例尺", FieldType.Integer);
                    GisTool.AddField(fc, "备注", FieldType.String);

                    pw.AddMessageMiddle(20, $"获取分组列表,并写入标记字段");

                    // 记录超限小班数量
                    int nullCount = 0;

                    // 获取行政村列表
                    List<string> names = fc.GetFieldValues(field_xzq);

                    foreach (string name in names)
                    {
                        
                        
                        if (name == "")    // 空值就跳过
                        {
                            continue;
                        }

                        // 编组号
                        int group = gp;

                        // 总行数
                        int rowCount = GetCount(fc, field_xzq, name);
                        
                        // 标记集合
                        Dictionary<int, Geometrys> dic = new Dictionary<int, Geometrys>();
                        // 起始号
                        int startRow = 1;

                        for (int k = 0; k < rowCount; k++)
                        {
                            // 超过总行数就退出
                            if (startRow > rowCount)
                            {
                                break;
                            }

                            if (startRow == k + 1)            // 起始号配对上
                            {
                                for (int i = 0; i < count; i++)
                                {
                                    // 当前集合数
                                    int IDCount = count - i;
                                    Geometrys geo = GetGeometrys(fc, field_xzq, name, IDCount, field_bh, startRow);
                                    // 获取要素的边界
                                    Envelope env = GeometryEngine.Instance.Boundary(geo.geometries).Extent;
                                    // 判断一下，如果地图框容得下，就返回编组号
                                    if (scWidth > env.Width && scHeight > env.Height)
                                    {
                                        geo.Scale = scale;
                                        dic.Add(group, geo);
                                        group++;
                                        startRow += geo.IDs.Count;
                                        break;
                                    }
                                    else
                                    {
                                        if (IDCount == 1)
                                        {
                                            geo.Scale = 0;
                                            dic.Add(group, geo);
                                            group++;
                                            startRow += geo.IDs.Count;
                                            nullCount++;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        // 写入标记
                        foreach (var item in dic)
                        {
                            int gp = item.Key;
                            Geometrys geo = item.Value;

                            Table table = fc.TargetTable();
                            string sql = $"{field_xzq} = '{name}' ";
                            QueryFilter queryFilter = new QueryFilter { WhereClause = sql };
                            using RowCursor rowCursor = table.Search(queryFilter, false);
                            while (rowCursor.MoveNext())
                            {
                                using Row row = rowCursor.Current;
                                var bhValue = row[field_bh];
                                if (bhValue is null)
                                {
                                    continue;
                                }
                                long bh = long.Parse(bhValue.ToString());

                                if (geo.IDs.Contains(bh))
                                {
                                    row["分组"] = gp;
                                    row["比例尺"] = geo.Scale;
                                    row["备注"] = geo.Scale == 0 ? "超限小班" : "";
                                    row.Store();
                                }
                            }
                        }
                    }

                    pw.AddMessageMiddle(50, $"链接回原图层，清理中间数据");

                    Arcpy.JoinField(fc_path, oidField, fc, "连接", new List<string>() { "分组", "比例尺", "备注" });
                    Arcpy.DeleteField(fc_path, "连接");
                    Arcpy.Delect(fc);

                    // 超限小班输出
                    if (nullCount>0)
                    {
                        pw.AddMessageMiddle(0, $"存在{nullCount}个超限小班！", Brushes.Red);
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


        // 收集Geometrys和IDs
        private Geometrys GetGeometrys(string fc, string field_xzq, string name, int IDCount, string field_bh, int startRow)
        {
            Geometrys geo = new Geometrys();
            // ID集合
            List<long> IDs = new List<long>();
            List<Geometry> geometries = new List<Geometry>();
            // 收集Geometry
            int index = 0;
            Table table = fc.TargetTable();
            string sql = $"{field_xzq} = '{name}' ";
            QueryFilter queryFilter = new QueryFilter { WhereClause = sql };
            using RowCursor rowCursor = table.Search(queryFilter, false);
            int initRow = 1;
            while (rowCursor.MoveNext())
            {
                // 跳过已处理的行
                if (initRow < startRow)
                {
                    initRow++;
                    continue;
                }
                // 进入正题
                if (index < IDCount)
                {
                    using Feature feature = rowCursor.Current as Feature;
                    long bh = long.Parse(feature[field_bh].ToString());
                    Geometry geometry = feature.GetShape();
                    // 加入集合
                    IDs.Add(bh);
                    geometries.Add(geometry);
                }
                index++;
            }

            Geometry unionGeo = GeometryEngine.Instance.Union(geometries);

            geo.IDs = IDs;
            geo.geometries = unionGeo;

            return geo;
        }


        // 获取总行数
        private int GetCount(string fc, string field_xzq, string name)
        {
            int rowCount = 0;
            Table table = fc.TargetTable();
            string sql = $"{field_xzq} = '{name}' ";
            QueryFilter queryFilter = new QueryFilter { WhereClause = sql };
            using RowCursor rowCursor = table.Search(queryFilter, false);
            while (rowCursor.MoveNext())
            {
                rowCount++;
            }

            return rowCount;
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/143854922";
            UITool.Link2Web(url);
        }
    }



    // Geometrys
    public class Geometrys
    {
        public List<long> IDs { get; set; }
        public Geometry geometries { get; set; }
        public int Scale { get; set; }
    }
}
