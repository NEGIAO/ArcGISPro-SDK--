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
    internal class ShowCalTFH : Button
    {

        private CalTFH _caltfh = null;

        protected override void OnClick()
        {
            //already open?
            if (_caltfh != null)
                return;
            _caltfh = new CalTFH();
            _caltfh.Owner = FrameworkApplication.Current.MainWindow;
            _caltfh.Closed += (o, e) => { _caltfh = null; };
            _caltfh.Show();
            //uncomment for modal
            //_caltfh.ShowDialog();
        }

    }
}
