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
    internal class ShowStatisticsYDYHD : Button
    {

        private StatisticsYDYHD _statisticsydyhd = null;

        protected override void OnClick()
        {
            //already open?
            if (_statisticsydyhd != null)
                return;
            _statisticsydyhd = new StatisticsYDYHD();
            _statisticsydyhd.Owner = FrameworkApplication.Current.MainWindow;
            _statisticsydyhd.Closed += (o, e) => { _statisticsydyhd = null; };
            _statisticsydyhd.Show();
            //uncomment for modal
            //_statisticsydyhd.ShowDialog();
        }

    }
}
