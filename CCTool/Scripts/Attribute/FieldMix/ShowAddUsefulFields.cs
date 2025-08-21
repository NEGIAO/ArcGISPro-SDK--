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

namespace CCTool.Scripts.Attribute.FieldMix
{
    internal class ShowAddUsefulFields : Button
    {

        private AddUsefulFields _addusefulfields = null;

        protected override void OnClick()
        {
            //already open?
            if (_addusefulfields != null)
                return;
            _addusefulfields = new AddUsefulFields();
            _addusefulfields.Owner = FrameworkApplication.Current.MainWindow;
            _addusefulfields.Closed += (o, e) => { _addusefulfields = null; };
            _addusefulfields.Show();
            //uncomment for modal
            //_addusefulfields.ShowDialog();
        }

    }
}
