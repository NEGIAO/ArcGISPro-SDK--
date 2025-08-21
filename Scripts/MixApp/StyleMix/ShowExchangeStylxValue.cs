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

namespace CCTool.Scripts.MixApp.StyleMix
{
    internal class ShowExchangeStylxValue : Button
    {

        private ExchangeStylxValue _exchangestylxvalue = null;

        protected override void OnClick()
        {
            //already open?
            if (_exchangestylxvalue != null)
                return;
            _exchangestylxvalue = new ExchangeStylxValue();
            _exchangestylxvalue.Owner = FrameworkApplication.Current.MainWindow;
            _exchangestylxvalue.Closed += (o, e) => { _exchangestylxvalue = null; };
            _exchangestylxvalue.Show();
            //uncomment for modal
            //_exchangestylxvalue.ShowDialog();
        }

    }
}
