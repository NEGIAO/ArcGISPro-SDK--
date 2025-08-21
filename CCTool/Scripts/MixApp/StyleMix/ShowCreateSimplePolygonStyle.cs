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

namespace CCTool.Scripts.MixApp.StyleMix
{
    internal class ShowCreateSimplePolygonStyle : Button
    {

        private CreateSimplePolygonStyle _createsimplepolygonstyle = null;

        protected override void OnClick()
        {
            //already open?
            if (_createsimplepolygonstyle != null)
                return;
            _createsimplepolygonstyle = new CreateSimplePolygonStyle();
            _createsimplepolygonstyle.Owner = FrameworkApplication.Current.MainWindow;
            _createsimplepolygonstyle.Closed += (o, e) => { _createsimplepolygonstyle = null; };
            _createsimplepolygonstyle.Show();
            //uncomment for modal
            //_createsimplepolygonstyle.ShowDialog();
        }

    }
}
