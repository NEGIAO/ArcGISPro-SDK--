using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using NPOI.POIFS.Crypt.Dsig;
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

namespace CCTool.Scripts.CusTool
{
    /// <summary>
    /// Interaction logic for SearchSameField.xaml
    /// </summary>
    public partial class SearchSameField : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public SearchSameField()
        {
            InitializeComponent();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "搜索相同字段值的图斑";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc_path = combox_fc.ComboxText();
                string fc_field_from = combox_field_from.ComboxText();
                string fc_field_bz = combox_field_bz.ComboxText();
                double dic = double.Parse(textDis.Text);

                // 默认数据库位置
                var gdb_path = Project.Current.DefaultGeodatabasePath;
                // 工程默认文件夹位置
                string folder_path = Project.Current.HomeFolderPath;

                // 判断参数是否选择完全
                if (fc_path == "" || fc_field_from == "" || fc_field_bz == "")
                {
                    MessageBox.Show("有必选参数为空或输入错误！！！");
                    return;
                }

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    FeatureLayer originFeatureLayer = fc_path.TargetFeatureLayer();
                    // 获取原始图层和标识图层的要素类
                    FeatureClass originFeatureClass = fc_path.TargetFeatureClass();

                    pw.AddMessageStart($"遍历要素");
                    // 获取目标图层和源图层的要素游标
                    using RowCursor originCursor = originFeatureClass.Search();

                    long index = 1;   // 计数器

                    // 获取OID字段
                    string oid = fc_path.TargetIDFieldName();
                    // 获取标记字段的长度
                    int bzLength = GisTool.GetFieldFromString(fc_path, fc_field_bz).Length;

                    // 遍历源图层的要素
                    while (originCursor.MoveNext())
                    {
                        // 计数标志
                        if (index % 200 == 0)
                        {
                            pw.AddMessageMiddle(0, $"累计图斑数量：{index}", Brushes.Gray);
                        }

                        // 标记文本
                        string bz = "";

                        using Feature originFeature = (Feature)originCursor.Current;
                        // 获取源要素的几何
                        ArcGIS.Core.Geometry.Geometry originGeometry = originFeature.GetShape();
                        // 用来空间分析的Geometry，主要看是否要buffer
                        ArcGIS.Core.Geometry.Geometry originGeometryBuffer = originGeometry;
                        originGeometryBuffer = GeometryEngine.Instance.Buffer(originGeometry, dic);

                        // 获取标记的字段值
                        var originFrom = originFeature[fc_field_from];
                        if (originFrom is null) { continue; }
                        string originFromValue = originFrom.ToString();

                        // OID值
                        string oidOrigin = originFeature[oid].ToString();

                        // 创建空间查询过滤器，以获取与源要素有重叠的目标要素
                        SpatialQueryFilter spatialFilter = new SpatialQueryFilter
                        {
                            FilterGeometry = originGeometryBuffer,
                            SpatialRelationship = SpatialRelationship.Intersects
                        };

                        // 在目标图层中查询与源要素重叠的要素
                        using RowCursor identityCursor = originFeatureClass.Search(spatialFilter);
                        while (identityCursor.MoveNext())
                        {
                            using Feature identityFeature = (Feature)identityCursor.Current;
                            // 获取目标要素的几何
                            ArcGIS.Core.Geometry.Geometry identityGeometry = identityFeature.GetShape();

                            // 计算源要素与目标要素的重叠面积
                            ArcGIS.Core.Geometry.Geometry intersection = GeometryEngine.Instance.Intersection(identityGeometry, originGeometryBuffer);
                            // 如果存在相交，则进行下一步处理
                            if (intersection != null)
                            {
                                // 获取标记的字段值
                                var identityFrom = identityFeature[fc_field_from];
                                if (identityFrom is null) { continue; }
                                string identityFromValue = identityFrom.ToString();
                                // OID值
                                string oidIdentity = identityFeature[oid].ToString();

                                if (identityFromValue == originFromValue && oidOrigin!=oidIdentity)   // 不是自身相交,且搜索值相同的情况
                                {
                                    // 更新标记值
                                    bz += $"{oidIdentity};";

                                }

                            }
                        }
                        // 赋值
                        if (bz.Length > bzLength)  // 如果字段太长，就截断
                        {
                            originFeature[fc_field_bz] = bz[..bzLength];
                        }
                        else
                        {
                            originFeature[fc_field_bz] = bz;
                        }

                        originFeature.Store();

                        index++;    // 计数加1
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
            string url = "https://blog.csdn.net/xcc34452366/article/details/139951117";
            UITool.Link2Web(url);
        }

        private void combox_field_from_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field_from);
        }

        private void combox_field_bz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_field_bz);
        }
    }
}
