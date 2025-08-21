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
    internal class ShowResort : Button
    {

        private Resort _resort = null;

        protected override void OnClick()
        {
            //already open?
            if (_resort != null)
                return;
            _resort = new Resort();
            _resort.Owner = FrameworkApplication.Current.MainWindow;
            _resort.Closed += (o, e) => { _resort = null; };
            _resort.Show();
            //uncomment for modal
            //_resort.ShowDialog();
        }

    }
}
