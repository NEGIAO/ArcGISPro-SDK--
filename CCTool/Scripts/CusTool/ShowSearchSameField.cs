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

namespace CCTool.Scripts.CusTool
{
    internal class ShowSearchSameField : Button
    {

        private SearchSameField _searchsamefield = null;

        protected override void OnClick()
        {
            //already open?
            if (_searchsamefield != null)
                return;
            _searchsamefield = new SearchSameField();
            _searchsamefield.Owner = FrameworkApplication.Current.MainWindow;
            _searchsamefield.Closed += (o, e) => { _searchsamefield = null; };
            _searchsamefield.Show();
            //uncomment for modal
            //_searchsamefield.ShowDialog();
        }

    }
}
