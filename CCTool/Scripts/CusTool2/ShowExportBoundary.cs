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

namespace CCTool.Scripts.CusTool2
{
    internal class ShowExportBoundary : Button
    {

        private ExportBoundary _exportboundary = null;

        protected override void OnClick()
        {
            //already open?
            if (_exportboundary != null)
                return;
            _exportboundary = new ExportBoundary();
            _exportboundary.Owner = FrameworkApplication.Current.MainWindow;
            _exportboundary.Closed += (o, e) => { _exportboundary = null; };
            _exportboundary.Show();
            //uncomment for modal
            //_exportboundary.ShowDialog();
        }

    }
}
