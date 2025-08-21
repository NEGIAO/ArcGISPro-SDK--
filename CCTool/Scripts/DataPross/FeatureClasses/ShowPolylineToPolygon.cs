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

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    internal class ShowPolylineToPolygon : Button
    {

        private PolylineToPolygon _polylinetopolygon = null;

        protected override void OnClick()
        {
            //already open?
            if (_polylinetopolygon != null)
                return;
            _polylinetopolygon = new PolylineToPolygon();
            _polylinetopolygon.Owner = FrameworkApplication.Current.MainWindow;
            _polylinetopolygon.Closed += (o, e) => { _polylinetopolygon = null; };
            _polylinetopolygon.Show();
            //uncomment for modal
            //_polylinetopolygon.ShowDialog();
        }

    }
}
