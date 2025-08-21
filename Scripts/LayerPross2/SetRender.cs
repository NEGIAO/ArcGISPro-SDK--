using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Extensions;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Dialogs;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.ToolManagers.Library;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.LayerPross2
{
    internal class SetRender : Button
    {
        protected override async void OnClick()
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    Map map = MapView.Active.Map;
                    // 获取图层
                    var featureLayers = MapView.Active.GetSelectedLayers().OfType<FeatureLayer>();

                    foreach (FeatureLayer featureLayer in featureLayers)
                    {
                        if (featureLayer is not null)
                        {
                            featureLayer.SetRenderer(GlobalData.renderer);
                        }
                    }
                });
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }
    }
}
