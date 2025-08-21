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
    internal class ShowLinkPolyline : Button
    {

        private LinkPolyline _linkpolyline = null;

        protected override void OnClick()
        {
            //already open?
            if (_linkpolyline != null)
                return;
            _linkpolyline = new LinkPolyline();
            _linkpolyline.Owner = FrameworkApplication.Current.MainWindow;
            _linkpolyline.Closed += (o, e) => { _linkpolyline = null; };
            _linkpolyline.Show();
            //uncomment for modal
            //_linkpolyline.ShowDialog();
        }

    }
}
