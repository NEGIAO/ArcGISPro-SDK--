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
    internal class ShowSortStylxItem : Button
    {

        private SortStylxItem _sortstylxitem = null;

        protected override void OnClick()
        {
            //already open?
            if (_sortstylxitem != null)
                return;
            _sortstylxitem = new SortStylxItem();
            _sortstylxitem.Owner = FrameworkApplication.Current.MainWindow;
            _sortstylxitem.Closed += (o, e) => { _sortstylxitem = null; };
            _sortstylxitem.Show();
            //uncomment for modal
            //_sortstylxitem.ShowDialog();
        }

    }
}
