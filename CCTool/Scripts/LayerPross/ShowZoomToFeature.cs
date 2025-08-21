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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.LayerPross
{
    internal class ShowZoomToFeature : Button
    {

        private ZoomToFeature _zoomtofeature = null;

        protected override void OnClick()
        {
            //already open?
            if (_zoomtofeature != null)
                return;
            _zoomtofeature = new ZoomToFeature();
            _zoomtofeature.Owner = FrameworkApplication.Current.MainWindow;
            _zoomtofeature.Closed += (o, e) => { _zoomtofeature = null; };
            _zoomtofeature.Show();
            //uncomment for modal
            //_zoomtofeature.ShowDialog();
        }

    }
}
