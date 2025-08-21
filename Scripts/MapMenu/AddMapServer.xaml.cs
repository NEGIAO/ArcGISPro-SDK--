using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Mapping.Ribbon;
using ArcGIS.Desktop.Mapping;
using Aspose.Words.Drawing;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Managers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
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

namespace CCTool.Scripts.MapMenu
{
    /// <summary>
    /// Interaction logic for AddMapServer.xaml
    /// </summary>
    public partial class AddMapServer : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public AddMapServer()
        {
            InitializeComponent();
            // 绑定TreeView数据
            this.tv.DataContext = MapItems.getData();

            // 初始化预设
            string mapName = BaseTool.ReadValueFromReg("mapSet", "name");
            string mapSource = BaseTool.ReadValueFromReg("mapSet", "source");

            textSource.Text = mapName;
            UITool.UpdataImageSource(IndexIamge, mapSource);

        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取参数
                string mapName = textSource.Text;

                // 创建临时文件夹
                string defFolder = Project.Current.HomeFolderPath;
                string targetFolder = $@"{defFolder}\CC配置文件(勿删)\在线影像";
                if (!Directory.Exists(targetFolder))
                {
                    Directory.CreateDirectory(targetFolder);
                }
                // 复制影像图层
                string oriPath = $@"CCTool.Data.Layers.ImageLayer.{mapName}.lyrx";
                string targetPath = $@"{targetFolder}\{mapName}.lyrx";
                DirTool.CopyResourceFile(oriPath, targetPath);

                Close();

                await QueuedTask.Run(() =>
                {
                    // 添加影像图层
                    MapCtlTool.AddLayerToMap(targetPath);
                    // 移到最底层
                    Map map = MapView.Active.Map;
                    Layer ly = map.FindLayers(mapName).FirstOrDefault();
                    map.MoveLayer(ly, -1);
                });
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void tv_SelectItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            try
            {
                MapItem mapItem = (MapItem)tv.SelectedItem;

                if (mapItem != null && mapItem.Width!=0)
                {
                    // 更新面板
                    textSource.Text = mapItem.MapName;
                    string source = @$"/CCTool;component/Data/IndexIamges/{mapItem.MapName}.png";
                    UITool.UpdataImageSource(IndexIamge, source);
                    // 写入预设
                    BaseTool.WriteValueToReg("mapSet", "name", textSource.Text);
                    BaseTool.WriteValueToReg("mapSet", "source", source);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
                return;
            }

        }

    }


    public class MapItem
    {
        public string MapName { get; set; }
        public string ImageSource { get; set; }
        public int Width { get; set; }
        public MapItems mapItems { get; set; }

        public void SetMapValue(string imagePath, List<string> mapNames)
        {
            this.mapItems = new MapItems();

            foreach (string mapName in mapNames)
            {
                this.mapItems.Add(new MapItem() { MapName = mapName, ImageSource = imagePath, Width = 12 });
            }
        }
    }

    public class MapItems : ObservableCollection<MapItem>
    {
        public static MapItems getData()
        {
            MapItems mapItems = new MapItems();

            // 影像底图
            MapItem mapItem = new MapItem() { MapName = "影像底图", Width = 0 };
            string imagePath = "/CCTool;component/Data/Icons/layer2.png";
            List<string> mapNames = new List<string>()
            {
                "Google影像","MapBox影像","天地图-影像地图（球面墨卡托投影）","2024040福建天地图",
                "高德地图","星图地球","World Imagery (Wayback 2024-11-18)",
            };
            mapItem.SetMapValue(imagePath, mapNames);
            mapItems.Add(mapItem);

            // 标注图
            mapItem = new MapItem() { MapName = "标注图", Width = 0 };
            imagePath = "/CCTool;component/Data/Icons/addPage.png";
            mapNames = new List<string>()
            {
                "天地图-影像注记（球面墨卡托投影）","天地图-地形注记（球面墨卡托投影）","天地图-矢量注记（球面墨卡托投影）",
            };
            mapItem.SetMapValue(imagePath, mapNames);
            mapItems.Add(mapItem);

            // 特殊主题地图
            mapItem = new MapItem() { MapName = "特殊主题地图", Width = 0 };
            imagePath = "/CCTool;component/Data/Icons/star.png";
            mapNames = new List<string>()
            {
                "世界海洋底图","地形底图","天地图-地形地图（球面墨卡托投影）","天地图-矢量地图（球面墨卡托投影）",
            };
            mapItem.SetMapValue(imagePath, mapNames);
            mapItems.Add(mapItem);

            return mapItems;
        }
    }
}
