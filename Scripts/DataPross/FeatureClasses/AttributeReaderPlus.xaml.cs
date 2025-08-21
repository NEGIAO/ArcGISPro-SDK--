using ActiproSoftware.Windows.Controls.DataGrid;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.MiniTool.GetInfo;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Windows;
using CCTool.Scripts.UI.ProMapTool;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    /// <summary>
    /// Interaction logic for AttributeReaderPlus.xaml
    /// </summary>
    public partial class AttributeReaderPlus : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public AttributeReaderPlus()
        {
            InitializeComponent();
        }

        //写在window里面，用来被外面的DataGrid绑定
        public ObservableCollection<string> m_associatedConditionsList
        {
            get { return _associatedConditionsList; }
            set { _associatedConditionsList = value; }
        }
        private ObservableCollection<string> _associatedConditionsList = new ObservableCollection<string>();


        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "属性读取加强版";

        private void combox_origin_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_origin_fc);
        }

        private void combox_identity_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_identity_fc);
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/139951330";
            UITool.Link2Web(url);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取数据
                string origin_fc = combox_origin_fc.ComboxText();
                string identity_fc = combox_identity_fc.ComboxText();

                string defGDB = Project.Current.DefaultGeodatabasePath;
                // 覆盖比例
                double prop = double.Parse(propTXT.Text) / 100;

                bool isWrite = (bool)isCopyNull.IsChecked;

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);


                // 获取对应字段列表
                List<List<string>> fieldPairs = new List<List<string>>();

                for (int i = 0; i < dg.Items.Count; i++)
                {
                    // 获取选择框
                    CheckBox item_isCheck = (CheckBox)dg.GetCell(i, 0).Content;
                    bool isCheck = (bool)item_isCheck.IsChecked;
                    // 获取目标字段
                    TextBlock item_originField = (TextBlock)dg.GetCell(i, 1).Content;
                    string originField = item_originField.Text;
                    // 获取来源字段
                    ComboBox item_identityField = (ComboBox)dg.GetCell(i, 4).Content;
                    string identityField = item_identityField.Text;
                    // 如果是选中状态，就加到列表中
                    if (isCheck)
                    {
                        fieldPairs.Add(new List<string>() { originField, identityField });
                    }
                }

                pw.AddMessageStart("检查数据");

                foreach (var field in fieldPairs)
                {
                    pw.AddMessageMiddle(0, $"目标字段：{field[0]}，来源字段：{field[1]}");
                    if (field[1] == "" || field[1] == null)   // 如果没有选择来源字段，中断
                    {
                        pw.AddMessageMiddle(0, $"目标字段{field[0]}没有选择相应的来源字段!", Brushes.Red);
                        return;
                    }
                }

                // 判断参数是否选择完全
                if (origin_fc == "" || identity_fc == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }


                Close();

                await QueuedTask.Run(() =>
                {
                    List<string> lines = new List<string>() { origin_fc, identity_fc };
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

                    // 获取原始图层和标识图层
                    FeatureLayer originFeatureLayer = origin_fc.TargetFeatureLayer();
                    FeatureLayer identityFeatureLayer = identity_fc.TargetFeatureLayer();
                    // 获取原始图层和标识图层的要素类
                    FeatureClass originFeatureClass = origin_fc.TargetFeatureClass();
                    FeatureClass identityFeatureClass = identity_fc.TargetFeatureClass();
                    // 如果两个图层使用的空间参考不同，投影到一致
                    SpatialReference sr_ori = originFeatureLayer.GetSpatialReference();
                    SpatialReference sr_iden = identityFeatureLayer.GetSpatialReference();
                    if (!sr_ori.Equals(sr_iden))
                    {
                        string MidFC = $@"{defGDB}\MidFC";
                        Arcpy.Project(identity_fc, $@"{defGDB}\MidFC", sr_ori.Wkt, sr_iden.Wkt);
                        identityFeatureClass = MidFC.TargetFeatureClass();
                    }

                    // 获取目标图层和源图层的要素游标
                    using RowCursor originCursor = originFeatureClass.Search();

                    long index = 1;   // 计数器
                    long featureCount = originFeatureLayer.GetFeatureCount();   // 目标图斑总数

                    // 遍历源图层的要素
                    while (originCursor.MoveNext())
                    {
                        // 计数标志
                        if (index % 200 == 0)
                        {
                            double process = 200 / (double)featureCount * 80;
                            pw.AddMessageMiddle(process, @$"累计图斑数量：{index}/{featureCount}", Brushes.Gray);
                        }

                        using Feature originFeature = (Feature)originCursor.Current;
                        double maxOverlapArea = 0;

                        // 初始化要标记的字段值
                        List<string> valueList = new List<string>();

                        Feature identityFeatureWithMaxOverlap = null;

                        // 获取源要素的几何
                        ArcGIS.Core.Geometry.Geometry originGeometry = originFeature.GetShape();

                        // 创建空间查询过滤器，以获取与源要素有重叠的目标要素
                        SpatialQueryFilter spatialFilter = new SpatialQueryFilter
                        {
                            FilterGeometry = originGeometry,
                            SpatialRelationship = SpatialRelationship.Intersects
                        };

                        // 在目标图层中查询与源要素重叠的要素
                        using (RowCursor identityCursor = identityFeatureClass.Search(spatialFilter))
                        {
                            while (identityCursor.MoveNext())
                            {
                                using Feature identityFeature = (Feature)identityCursor.Current;
                                // 获取目标要素的几何
                                ArcGIS.Core.Geometry.Geometry identityGeometry = identityFeature.GetShape();

                                // 计算源要素与目标要素的重叠面积
                                ArcGIS.Core.Geometry.Geometry intersection = GeometryEngine.Instance.Intersection(originGeometry, identityGeometry);
                                double overlapArea = Math.Round((intersection as ArcGIS.Core.Geometry.Polygon).Area, 2);
                                double originArea = Math.Round((originGeometry as ArcGIS.Core.Geometry.Polygon).Area, 2);

                                // 如果重叠面积大于当前最大重叠面积，则更新最大重叠面积和目标要素
                                if (overlapArea > maxOverlapArea && overlapArea / originArea >= prop)
                                {
                                    maxOverlapArea = overlapArea;
                                    // 重叠Feature
                                    identityFeatureWithMaxOverlap = identityFeature;

                                    // 赋值
                                    valueList.Clear();
                                    foreach (var fieldPair in fieldPairs)
                                    {
                                        valueList.Add(identityFeature[fieldPair[1]]?.ToString());
                                    }
                                }
                            }
                        }

                        // 如果找到与源要素有最大重叠的目标要素，则将其属性复制到源要素
                        if (identityFeatureWithMaxOverlap != null)
                        {
                            // 复制属性
                            if (isWrite)
                            {
                                for (int i = 0; i < fieldPairs.Count; i++)
                                {
                                    string targetValue = originFeature[fieldPairs[i][0]]?.ToString();
                                    if (targetValue == "" || targetValue is null)
                                    {
                                        originFeature[fieldPairs[i][0]] = valueList[i];
                                    }
                                }
                            }
                            else
                            {
                                for (int i = 0; i < fieldPairs.Count; i++)
                                {

                                    originFeature[fieldPairs[i][0]] = valueList[i];
                                }

                            }
                        }
                        else    // 不符合要求的情况下
                        {
                            //// 清空值
                            //foreach (List<string> fieldPair in fieldPairs)
                            //{
                            //    originFeature[fieldPair[0]] = null;
                            //}
                        }
                        // 更新源图层中的源要素
                        originFeature.Store();

                        index++;
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


        // 定义一个空包
        List<FieldAtt> fieldAtt = new List<FieldAtt>();

        private async void combox_origin_fc_Closed(object sender, EventArgs e)
        {
            try
            {
                // 定义一个空包
                List<FieldAtt> fieldAtt2 = new List<FieldAtt>();

                // 将字段属性加入到dg中
                string lyName = combox_origin_fc.ComboxText();

                if (lyName == "") { return; }

                var fields = await QueuedTask.Run(() =>
                {
                    return GisTool.GetFieldsFromTarget(lyName, "all");
                });
                // 添加数据
                foreach (Field field in fields)
                {
                    fieldAtt2.Add(new FieldAtt()
                    {
                        FieldName = field.Name,
                        FieldType = field.FieldType.ToString(),
                        FieldLength = field.Length.ToString(),
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

        // 选择来源图层后，更新来源字段
        private async void combox_identity_fc_Closed(object sender, EventArgs e)
        {
            try
            {
                // 清空来源字段列表
                _associatedConditionsList.Clear();

                // 将字段属性加入到dg中
                string lyName = combox_identity_fc.ComboxText();

                if (lyName == "") { return; }

                var fields = await QueuedTask.Run(() =>
                {
                    return GisTool.GetFieldsFromTarget(lyName, "all");
                });
                // 添加数据
                foreach (Field field in fields)
                {
                    _associatedConditionsList.Add(field.Name);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private List<string> CheckData(List<string> lines)
        {
            List<string> result = new List<string>();


            return result;
        }

    }

    //写在Window外面，总之要保证被local:到
    public class ProtectInfo : INotifyPropertyChanged
    {
        public ProtectInfo(string associatedConditions)
        {
            _associatedConditions = associatedConditions;
        }
        private string _associatedConditions;
        public string AssociatedConditions
        {
            get { return _associatedConditions; }
            set
            {
                if (_associatedConditions != value)
                {
                    _associatedConditions = value;
                    OnPropertyChanged(nameof(_associatedConditions));
                }
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }


    class FieldAtt
    {
        public bool IsCheck { get; set; }
        public string FieldName { get; set; }
        public string FieldType { get; set; }
        public string FieldLength { get; set; }
    }

}
