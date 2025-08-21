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

namespace CCTool.Scripts.GTApp.ZYR
{
    internal class ShowZY2 : Button
    {

        private ZY2 _zy2 = null;

        protected override void OnClick()
        {
            //already open?
            if (_zy2 != null)
                return;
            _zy2 = new ZY2();
            _zy2.Owner = FrameworkApplication.Current.MainWindow;
            _zy2.Closed += (o, e) => { _zy2 = null; };
            _zy2.Show();
            //uncomment for modal
            //_zy2.ShowDialog();
        }

    }
}
