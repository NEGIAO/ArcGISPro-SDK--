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
    internal class ShowDat2FeatureClass : Button
    {

        private Dat2FeatureClass _dat2featureclass = null;

        protected override void OnClick()
        {
            //already open?
            if (_dat2featureclass != null)
                return;
            _dat2featureclass = new Dat2FeatureClass();
            _dat2featureclass.Owner = FrameworkApplication.Current.MainWindow;
            _dat2featureclass.Closed += (o, e) => { _dat2featureclass = null; };
            _dat2featureclass.Show();
            //uncomment for modal
            //_dat2featureclass.ShowDialog();
        }

    }
}
