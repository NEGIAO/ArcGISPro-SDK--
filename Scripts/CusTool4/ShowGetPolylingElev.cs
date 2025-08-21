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

namespace CCTool.Scripts.CusTool4
{
    internal class ShowGetPolylingElev : Button
    {

        private GetPolylingElev _getpolylingelev = null;

        protected override void OnClick()
        {
            //already open?
            if (_getpolylingelev != null)
                return;
            _getpolylingelev = new GetPolylingElev();
            _getpolylingelev.Owner = FrameworkApplication.Current.MainWindow;
            _getpolylingelev.Closed += (o, e) => { _getpolylingelev = null; };
            _getpolylingelev.Show();
            //uncomment for modal
            //_getpolylingelev.ShowDialog();
        }

    }
}
