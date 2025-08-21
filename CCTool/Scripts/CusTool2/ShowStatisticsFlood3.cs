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
    internal class ShowStatisticsFlood3 : Button
    {

        private StatisticsFlood3 _statisticsflood3 = null;

        protected override void OnClick()
        {
            //already open?
            if (_statisticsflood3 != null)
                return;
            _statisticsflood3 = new StatisticsFlood3();
            _statisticsflood3.Owner = FrameworkApplication.Current.MainWindow;
            _statisticsflood3.Closed += (o, e) => { _statisticsflood3 = null; };
            _statisticsflood3.Show();
            //uncomment for modal
            //_statisticsflood3.ShowDialog();
        }

    }
}
