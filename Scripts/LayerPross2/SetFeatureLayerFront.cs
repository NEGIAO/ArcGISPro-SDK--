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
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.LayerPross2
{
    internal class SetFeatureLayerFront : Button
    {
        protected override async void OnClick()
        {
            await QueuedTask.Run(() =>
            {
                Map map = MapView.Active.Map;
                // 获取图层
                var lys = MapView.Active.GetSelectedLayers().ToList();
                // 倒排
                lys.Reverse();

                foreach (Layer ly in lys)
                {
                    map.MoveLayer(ly, 0);
                }

            });

        }
    }
}
