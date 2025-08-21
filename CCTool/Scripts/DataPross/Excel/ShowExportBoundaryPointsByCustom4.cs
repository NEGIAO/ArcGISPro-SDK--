using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

namespace CCTool.Scripts.DataPross.Excel
{
    internal class ShowExportBoundaryPointsByCustom4 : Button
{

    private ExportBoundaryPointsByCustom4 _exportboundarypointsbycustom4 = null;

    protected override void OnClick()
    {
        //already open?
        if (_exportboundarypointsbycustom4 != null)
            return;
        _exportboundarypointsbycustom4 = new ExportBoundaryPointsByCustom4();
        _exportboundarypointsbycustom4.Owner = FrameworkApplication.Current.MainWindow;
        _exportboundarypointsbycustom4.Closed += (o, e) => { _exportboundarypointsbycustom4 = null; };
        _exportboundarypointsbycustom4.Show();
         //uncomment for modal
         //_exportboundarypointsbycustom4.ShowDialog();
}

}
}
