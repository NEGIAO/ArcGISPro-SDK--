using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

namespace CCTool.Scripts.GDBMenu
{
    internal class ShowCalculateYSDM : Button
{

    private CalculateYSDM _calculateysdm = null;

    protected override void OnClick()
    {
        //already open?
        if (_calculateysdm != null)
            return;
        _calculateysdm = new CalculateYSDM();
        _calculateysdm.Owner = FrameworkApplication.Current.MainWindow;
        _calculateysdm.Closed += (o, e) => { _calculateysdm = null; };
        _calculateysdm.Show();
         //uncomment for modal
         //_calculateysdm.ShowDialog();
}

}
}
