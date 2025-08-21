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

namespace CCTool.Scripts.MapMenu
{
    internal class ShowAddMapServer : Button
    {

        private AddMapServer _addmapserver = null;

        protected override void OnClick()
        {
            //already open?
            if (_addmapserver != null)
                return;
            _addmapserver = new AddMapServer();
            _addmapserver.Owner = FrameworkApplication.Current.MainWindow;
            _addmapserver.Closed += (o, e) => { _addmapserver = null; };
            _addmapserver.Show();
            //uncomment for modal
            //_addmapserver.ShowDialog();
        }

    }
}
