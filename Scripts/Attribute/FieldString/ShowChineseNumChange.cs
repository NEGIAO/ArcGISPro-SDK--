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

namespace CCTool.Scripts.Attribute.FieldString
{
    internal class ShowChineseNumChange : Button
    {

        private ChineseNumChange _chinesenumchange = null;

        protected override void OnClick()
        {
            //already open?
            if (_chinesenumchange != null)
                return;
            _chinesenumchange = new ChineseNumChange();
            _chinesenumchange.Owner = FrameworkApplication.Current.MainWindow;
            _chinesenumchange.Closed += (o, e) => { _chinesenumchange = null; };
            _chinesenumchange.Show();
            //uncomment for modal
            //_chinesenumchange.ShowDialog();
        }

    }
}
