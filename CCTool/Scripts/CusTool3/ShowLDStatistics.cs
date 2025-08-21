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

namespace CCTool.Scripts.CusTool3
{
    internal class ShowLDStatistics : Button
    {

        private LDStatistics _ldstatistics = null;

        protected override void OnClick()
        {
            //already open?
            if (_ldstatistics != null)
                return;
            _ldstatistics = new LDStatistics();
            _ldstatistics.Owner = FrameworkApplication.Current.MainWindow;
            _ldstatistics.Closed += (o, e) => { _ldstatistics = null; };
            _ldstatistics.Show();
            //uncomment for modal
            //_ldstatistics.ShowDialog();
        }

    }
}
