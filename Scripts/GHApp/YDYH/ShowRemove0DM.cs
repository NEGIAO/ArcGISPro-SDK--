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

namespace CCTool.Scripts.GHApp.YDYH
{
    internal class ShowRemove0DM : Button
    {

        private Remove0DM _remove0dm = null;

        protected override void OnClick()
        {
            //already open?
            if (_remove0dm != null)
                return;
            _remove0dm = new Remove0DM();
            _remove0dm.Owner = FrameworkApplication.Current.MainWindow;
            _remove0dm.Closed += (o, e) => { _remove0dm = null; };
            _remove0dm.Show();
            //uncomment for modal
            //_remove0dm.ShowDialog();
        }

    }
}
