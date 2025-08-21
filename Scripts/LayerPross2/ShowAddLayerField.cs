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

namespace CCTool.Scripts.LayerPross2
{
    internal class ShowAddLayerField : Button
    {

        private AddLayerField _addlayerfield = null;

        protected override void OnClick()
        {
            //already open?
            if (_addlayerfield != null)
                return;
            _addlayerfield = new AddLayerField();
            _addlayerfield.Owner = FrameworkApplication.Current.MainWindow;
            _addlayerfield.Closed += (o, e) => { _addlayerfield = null; };
            _addlayerfield.Show();
            //uncomment for modal
            //_addlayerfield.ShowDialog();
        }

    }
}
