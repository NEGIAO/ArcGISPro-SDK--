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

namespace CCTool.Scripts.LayerExport
{
    internal class ShowExport2CADPlus : Button
    {

        private Export2CADPlus _export2cadplus = null;

        protected override void OnClick()
        {
            //already open?
            if (_export2cadplus != null)
                return;
            _export2cadplus = new Export2CADPlus();
            _export2cadplus.Owner = FrameworkApplication.Current.MainWindow;
            _export2cadplus.Closed += (o, e) => { _export2cadplus = null; };
            _export2cadplus.Show();
            //uncomment for modal
            //_export2cadplus.ShowDialog();
        }

    }
}
