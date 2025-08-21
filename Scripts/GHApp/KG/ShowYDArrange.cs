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

namespace CCTool.Scripts.GHApp.KG
{
    internal class ShowYDArrange : Button
    {

        private YDArrange _ydarrange = null;

        protected override void OnClick()
        {
            //already open?
            if (_ydarrange != null)
                return;
            _ydarrange = new YDArrange();
            _ydarrange.Owner = FrameworkApplication.Current.MainWindow;
            _ydarrange.Closed += (o, e) => { _ydarrange = null; };
            _ydarrange.Show();
            //uncomment for modal
            //_ydarrange.ShowDialog();
        }

    }
}
