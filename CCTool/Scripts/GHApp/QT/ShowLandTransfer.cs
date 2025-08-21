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

namespace CCTool.Scripts.GHApp.QT
{
    internal class ShowLandTransfer : Button
    {

        private LandTransfer _landtransfer = null;

        protected override void OnClick()
        {
            //already open?
            if (_landtransfer != null)
                return;
            _landtransfer = new LandTransfer();
            _landtransfer.Owner = FrameworkApplication.Current.MainWindow;
            _landtransfer.Closed += (o, e) => { _landtransfer = null; };
            _landtransfer.Show();
            //uncomment for modal
            //_landtransfer.ShowDialog();
        }

    }
}
