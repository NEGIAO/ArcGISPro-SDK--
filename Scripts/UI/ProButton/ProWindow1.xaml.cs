using ArcGIS.Desktop.Framework;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Enumeration;
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

namespace CCTool.Scripts.UI.ProButton
{
    /// <summary>
    /// Interaction logic for ProWindow1.xaml
    /// </summary>
    public partial class ProWindow1 : ArcGIS.Desktop.Framework.Controls.ProWindow
    {

        public ProWindow1()
        {
            InitializeComponent();

           

        }


    }


    public class Color
    {
        public string ImageName { get; set; }
        public string Name { get; set; }
    }

    public class Item
    {
        public string Name { get; set; }
        public string ImageName { get; set; }
        public int Width { get; set; }
        public Items Children { get; set; }


        string imagePath = "/CCTool;component/Data/Icons/layer2.png";
        

        public void SetChildrenDefalutValue()
        {
            this.Children = new Items();
            this.Children.Add(new Item() { Name = "子元素1" , ImageName= imagePath, Width =12 });
            this.Children.Add(new Item() { Name = "子元素2", ImageName = imagePath, Width = 12 });
            this.Children.Add(new Item() { Name = "子元素3", ImageName = imagePath, Width = 12 });
        }
    }

    public class Items : ObservableCollection<Item>
    {
        public static Items getTestData()
        {
            Items items = new Items();

            Item item = new Item() { Name = "影像底图", Width = 0 };
            item.SetChildrenDefalutValue();
            items.Add(item);

            item = new Item() { Name = "标注图", Width = 0 };
            item.SetChildrenDefalutValue();
            items.Add(item);

            item = new Item() { Name = "特殊主题地图", Width = 0 };
            item.SetChildrenDefalutValue();
            items.Add(item);

            return items;
        }
    }
}
