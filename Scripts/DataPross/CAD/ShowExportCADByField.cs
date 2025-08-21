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

namespace CCTool.Scripts.DataPross.CAD
{
    internal class ShowExportCADByField : Button
    {

        private ExportCADByField _exportcadbyfield = null;

        protected override void OnClick()
        {
            //already open?
            if (_exportcadbyfield != null)
                return;
            _exportcadbyfield = new ExportCADByField();
            _exportcadbyfield.Owner = FrameworkApplication.Current.MainWindow;
            _exportcadbyfield.Closed += (o, e) => { _exportcadbyfield = null; };
            _exportcadbyfield.Show();
            //uncomment for modal
            //_exportcadbyfield.ShowDialog();
        }

    }
}
