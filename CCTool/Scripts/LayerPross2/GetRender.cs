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
    internal class GetRender : Button
    {
        protected override async void OnClick()
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    Map map = MapView.Active.Map;
                    // 获取图层
                    FeatureLayer featureLayer = MapView.Active.GetSelectedLayers().OfType<FeatureLayer>().FirstOrDefault();

                    //  判定是不是要素图层
                    if (featureLayer is null)
                    {
                        MessageBox.Show("请选择一个要素图层！");
                    }
                    else 
                    { 
                        GlobalData.renderer = featureLayer.GetRenderer();
                    }

                });
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message+ee.StackTrace);
                return;
            }
            
        }
    }
}
