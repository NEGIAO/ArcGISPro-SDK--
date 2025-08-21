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
    internal class ShowZY3 : Button
    {

        private ZY3 _zy3 = null;

        protected override void OnClick()
        {
            //already open?
            if (_zy3 != null)
                return;
            _zy3 = new ZY3();
            _zy3.Owner = FrameworkApplication.Current.MainWindow;
            _zy3.Closed += (o, e) => { _zy3 = null; };
            _zy3.Show();
            //uncomment for modal
            //_zy3.ShowDialog();
        }

    }
}
