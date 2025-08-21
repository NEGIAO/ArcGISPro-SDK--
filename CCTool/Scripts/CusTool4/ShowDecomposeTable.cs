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

namespace CCTool.Scripts.CusTool4
{
    internal class ShowDecomposeTable : Button
    {

        private DecomposeTable _decomposetable = null;

        protected override void OnClick()
        {
            //already open?
            if (_decomposetable != null)
                return;
            _decomposetable = new DecomposeTable();
            _decomposetable.Owner = FrameworkApplication.Current.MainWindow;
            _decomposetable.Closed += (o, e) => { _decomposetable = null; };
            _decomposetable.Show();
            //uncomment for modal
            //_decomposetable.ShowDialog();
        }

    }
}
