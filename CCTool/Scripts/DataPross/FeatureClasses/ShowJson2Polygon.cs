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
    internal class ShowJson2Polygon : Button
    {

        private Json2Polygon _json2polygon = null;

        protected override void OnClick()
        {
            //already open?
            if (_json2polygon != null)
                return;
            _json2polygon = new Json2Polygon();
            _json2polygon.Owner = FrameworkApplication.Current.MainWindow;
            _json2polygon.Closed += (o, e) => { _json2polygon = null; };
            _json2polygon.Show();
            //uncomment for modal
            //_json2polygon.ShowDialog();
        }

    }
}
