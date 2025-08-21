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

namespace CCTool.Scripts.DataPross.FeatureClasses
{
    internal class ShowAdjustmentsTool2 : Button
    {

        private AdjustmentsTool2 _adjustmentstool2 = null;

        protected override void OnClick()
        {
            //already open?
            if (_adjustmentstool2 != null)
                return;
            _adjustmentstool2 = new AdjustmentsTool2();
            _adjustmentstool2.Owner = FrameworkApplication.Current.MainWindow;
            _adjustmentstool2.Closed += (o, e) => { _adjustmentstool2 = null; };
            _adjustmentstool2.Show();
            //uncomment for modal
            //_adjustmentstool2.ShowDialog();
        }

    }
}
