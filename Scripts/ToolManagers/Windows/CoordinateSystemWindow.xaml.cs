using ArcGIS.Core.Geometry;
using CCTool.Scripts.ToolManagers.Library;
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

namespace CCTool.Scripts.ToolManagers.Windows
{
    /// <summary>
    /// Interaction logic for CoordinateSystemWindow.xaml
    /// </summary>
    public partial class CoordinateSystemWindow : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public CoordinateSystemWindow()
        {
            InitializeComponent();
        }

        private void btn_go_Click(object sender, RoutedEventArgs e)
        {
            Close();

            GlobalData.sr = cod.SelectedSpatialReference;
            // 通知事件
            EventCenter.Broadcast(EventDefine.UpdataCod);
        }
    }
}
