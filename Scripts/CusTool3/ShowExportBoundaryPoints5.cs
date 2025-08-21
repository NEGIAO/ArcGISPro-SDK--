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

namespace CCTool.Scripts.CusTool3
{
    internal class ShowExportBoundaryPoints5 : Button
    {

        private ExportBoundaryPoints5 _exportboundarypoints5 = null;

        protected override void OnClick()
        {
            //already open?
            if (_exportboundarypoints5 != null)
                return;
            _exportboundarypoints5 = new ExportBoundaryPoints5();
            _exportboundarypoints5.Owner = FrameworkApplication.Current.MainWindow;
            _exportboundarypoints5.Closed += (o, e) => { _exportboundarypoints5 = null; };
            _exportboundarypoints5.Show();
            //uncomment for modal
            //_exportboundarypoints5.ShowDialog();
        }

    }
}
