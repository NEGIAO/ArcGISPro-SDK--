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

namespace CCTool.Scripts.GDBMenu
{
    internal class ShowClearGDBAll : Button
    {

        private ClearGDBAll _cleargdball = null;

        protected override void OnClick()
        {
            //already open?
            if (_cleargdball != null)
                return;
            _cleargdball = new ClearGDBAll();
            _cleargdball.Owner = FrameworkApplication.Current.MainWindow;
            _cleargdball.Closed += (o, e) => { _cleargdball = null; };
            _cleargdball.Show();
            //uncomment for modal
            //_cleargdball.ShowDialog();
        }

    }
}
