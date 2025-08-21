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

namespace CCTool.Scripts.MiniTool.MTool
{
    internal class ShowNum2Chinese : Button
    {

        private Num2Chinese _num2chinese = null;

        protected override void OnClick()
        {
            //already open?
            if (_num2chinese != null)
                return;
            _num2chinese = new Num2Chinese();
            _num2chinese.Owner = FrameworkApplication.Current.MainWindow;
            _num2chinese.Closed += (o, e) => { _num2chinese = null; };
            _num2chinese.Show();
            //uncomment for modal
            //_num2chinese.ShowDialog();
        }

    }
}
