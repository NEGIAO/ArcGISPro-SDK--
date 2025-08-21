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
    internal class ShowAttributeReaderPlus : Button
    {

        private AttributeReaderPlus _attributereaderplus = null;

        protected override void OnClick()
        {
            //already open?
            if (_attributereaderplus != null)
                return;
            _attributereaderplus = new AttributeReaderPlus();
            _attributereaderplus.Owner = FrameworkApplication.Current.MainWindow;
            _attributereaderplus.Closed += (o, e) => { _attributereaderplus = null; };
            _attributereaderplus.Show();
            //uncomment for modal
            //_attributereaderplus.ShowDialog();
        }

    }
}
