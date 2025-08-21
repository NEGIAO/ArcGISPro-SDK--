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

namespace CCTool.Scripts.CusTool4
{
    internal class ShowExportBoundarySingle : Button
    {

        private ExportBoundarySingle _exportboundarysingle = null;

        protected override void OnClick()
        {
            //already open?
            if (_exportboundarysingle != null)
                return;
            _exportboundarysingle = new ExportBoundarySingle();
            _exportboundarysingle.Owner = FrameworkApplication.Current.MainWindow;
            _exportboundarysingle.Closed += (o, e) => { _exportboundarysingle = null; };
            _exportboundarysingle.Show();
            //uncomment for modal
            //_exportboundarysingle.ShowDialog();
        }

    }
}
