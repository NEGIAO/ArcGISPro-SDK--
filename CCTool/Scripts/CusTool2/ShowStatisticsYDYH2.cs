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
    internal class ShowStatisticsYDYH2 : Button
    {

        private StatisticsYDYH2 _statisticsydyh2 = null;

        protected override void OnClick()
        {
            //already open?
            if (_statisticsydyh2 != null)
                return;
            _statisticsydyh2 = new StatisticsYDYH2();
            _statisticsydyh2.Owner = FrameworkApplication.Current.MainWindow;
            _statisticsydyh2.Closed += (o, e) => { _statisticsydyh2 = null; };
            _statisticsydyh2.Show();
            //uncomment for modal
            //_statisticsydyh2.ShowDialog();
        }

    }
}
