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
    internal class ShowYDAnalysis : Button
    {

        private YDAnalysis _ydanalysis = null;

        protected override void OnClick()
        {
            //already open?
            if (_ydanalysis != null)
                return;
            _ydanalysis = new YDAnalysis();
            _ydanalysis.Owner = FrameworkApplication.Current.MainWindow;
            _ydanalysis.Closed += (o, e) => { _ydanalysis = null; };
            _ydanalysis.Show();
            //uncomment for modal
            //_ydanalysis.ShowDialog();
        }

    }
}
