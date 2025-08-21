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
    internal class ShowSupply0DM : Button
    {

        private Supply0DM _supply0dm = null;

        protected override void OnClick()
        {
            //already open?
            if (_supply0dm != null)
                return;
            _supply0dm = new Supply0DM();
            _supply0dm.Owner = FrameworkApplication.Current.MainWindow;
            _supply0dm.Closed += (o, e) => { _supply0dm = null; };
            _supply0dm.Show();
            //uncomment for modal
            //_supply0dm.ShowDialog();
        }

    }
}
