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
    internal class ShowCalculateBSM : Button
    {

        private CalculateBSM _calculatebsm = null;

        protected override void OnClick()
        {
            //already open?
            if (_calculatebsm != null)
                return;
            _calculatebsm = new CalculateBSM();
            _calculatebsm.Owner = FrameworkApplication.Current.MainWindow;
            _calculatebsm.Closed += (o, e) => { _calculatebsm = null; };
            _calculatebsm.Show();
            //uncomment for modal
            //_calculatebsm.ShowDialog();
        }

    }
}
