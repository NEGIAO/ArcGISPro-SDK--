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
    internal class ShowExcel2PolygonCom : Button
    {

        private Excel2PolygonCom _excel2polygoncom = null;

        protected override void OnClick()
        {
            //already open?
            if (_excel2polygoncom != null)
                return;
            _excel2polygoncom = new Excel2PolygonCom();
            _excel2polygoncom.Owner = FrameworkApplication.Current.MainWindow;
            _excel2polygoncom.Closed += (o, e) => { _excel2polygoncom = null; };
            _excel2polygoncom.Show();
            //uncomment for modal
            //_excel2polygoncom.ShowDialog();
        }

    }
}
