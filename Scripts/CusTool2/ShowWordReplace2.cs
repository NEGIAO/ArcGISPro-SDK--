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
    internal class ShowWordReplace2 : Button
    {

        private WordReplace2 _wordreplace2 = null;

        protected override void OnClick()
        {
            //already open?
            if (_wordreplace2 != null)
                return;
            _wordreplace2 = new WordReplace2();
            _wordreplace2.Owner = FrameworkApplication.Current.MainWindow;
            _wordreplace2.Closed += (o, e) => { _wordreplace2 = null; };
            _wordreplace2.Show();
            //uncomment for modal
            //_wordreplace2.ShowDialog();
        }

    }
}
