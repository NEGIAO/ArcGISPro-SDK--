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

namespace CCTool.Scripts.MixApp.StyleMix
{
    internal class ShowStylxToFeatureLayer : Button
    {

        private StylxToFeatureLayer _stylxtofeaturelayer = null;

        protected override void OnClick()
        {
            //already open?
            if (_stylxtofeaturelayer != null)
                return;
            _stylxtofeaturelayer = new StylxToFeatureLayer();
            _stylxtofeaturelayer.Owner = FrameworkApplication.Current.MainWindow;
            _stylxtofeaturelayer.Closed += (o, e) => { _stylxtofeaturelayer = null; };
            _stylxtofeaturelayer.Show();
            //uncomment for modal
            //_stylxtofeaturelayer.ShowDialog();
        }

    }
}
