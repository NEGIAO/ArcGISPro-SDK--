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
    internal class ShowPhoto2Point : Button
    {

        private Photo2Point _photo2point = null;

        protected override void OnClick()
        {
            //already open?
            if (_photo2point != null)
                return;
            _photo2point = new Photo2Point();
            _photo2point.Owner = FrameworkApplication.Current.MainWindow;
            _photo2point.Closed += (o, e) => { _photo2point = null; };
            _photo2point.Show();
            //uncomment for modal
            //_photo2point.ShowDialog();
        }

    }
}
