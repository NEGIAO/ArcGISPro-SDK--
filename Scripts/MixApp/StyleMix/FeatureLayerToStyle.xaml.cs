using ArcGIS.Core.CIM;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Library;
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
using Path = System.IO.Path;

namespace CCTool.Scripts.MixApp.StyleMix
{
    /// <summary>
    /// Interaction logic for FeatureLayerToStyle.xaml
    /// </summary>
    public partial class FeatureLayerToStyle : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "FeatureLayerToStyle";

        public FeatureLayerToStyle()
        {
            InitializeComponent();

            // 初始化参数选项
            cb_add.IsChecked = BaseTool.ReadValueFromReg(toolSet, "isAdd").ToBool();

        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "图层唯一值符号转样式库";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc, "All");
        }

        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147043057";
            UITool.Link2Web(url);
        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string fc = combox_fc.ComboxText();
                string stylxName = combox_stylx.Text;
                bool isAdd = (bool)cb_add.IsChecked;

                // 判断参数是否选择完全
                if (fc == "" || stylxName == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }
                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "isAdd", isAdd);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("检查数据");
                    // 检查数据
                    List<string> errs = CheckData(fc);
                    // 打印错误
                    if (errs.Count > 0)
                    {
                        foreach (var err in errs)
                        {
                            pw.AddMessageMiddle(10, err, Brushes.Red);
                        }
                        return;
                    }

                    pw.AddMessageMiddle(10,"获取图层的唯一值符号");
                    // 获取图层
                    FeatureLayer featureLayer = fc.TargetFeatureLayer();

                    // 获取图层的唯一值符号的字典
                    Dictionary<string, StylxAtt> valueSymbolDict = StylxTool.GetFeatureLayerSymbolDict(featureLayer);

                    pw.AddMessageMiddle(20, "预设.stylx文件");
                    // 创建.stylx文件
                    CreateStylx(stylxName, isAdd);

                    pw.AddMessageMiddle(20, "写入.stylx文件");

                    StyleProjectItem styleProjectItem = stylxName.TargetStyleProjectItem();
                    // 写入.stylx文件
                    StylxTool.WriteStylxItem(styleProjectItem, valueSymbolDict);

                });
                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private List<string> CheckData(string fc)
        {
            List<string> result = new List<string>();

            //假设已有FeatureLayer实例
            FeatureLayer featureLayer = fc.TargetFeatureLayer();

            //获取唯一值渲染器
            var renderer = featureLayer.GetRenderer() as CIMUniqueValueRenderer;
            if (renderer == null) 
            {
                result.Add("目标图层的符号样式不是唯一值类型！") ;
            };

            return result;
        }


        private void CreateStylx(string stylxName, bool isAdd)
        {
            string defFolder = Project.Current.HomeFolderPath;
            // 新建的路径
            string defPath = $@"{defFolder}\{stylxName}.stylx";

            //获取当前工程中的所有tylx
            var ProjectStyles = Project.Current.GetItems<StyleProjectItem>();
            //根据名字找出指定的tylx
            StyleProjectItem style = ProjectStyles.FirstOrDefault(x => x.Name == stylxName);
            // 如果当前地图不存在该stylx
            if (style is null)
            {
                StyleHelper.CreateStyle(Project.Current, defPath);
            }
            // 如果已经有该stylx
            else
            {
                // stylx文件原始路径
                string stylePath = style.Path;
                // 如果不是追加，就删除并新建
                if (!isAdd)
                {
                    StyleHelper.RemoveStyle(Project.Current, stylePath);
                    File.Delete(stylePath);
                    StyleHelper.CreateStyle(Project.Current, stylePath);
                }
            }


        }

        private void combox_stylx_DropDown(object sender, EventArgs e)
        {
            UITool.AddStylxsToCombox(combox_stylx);
        }
    }
}
