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

namespace CCTool.Scripts.TableMenu
{
    internal class ShowFieldStatistics : Button
    {

        private FieldStatistics _fieldstatistics = null;

        protected override void OnClick()
        {
            //already open?
            if (_fieldstatistics != null)
                return;
            _fieldstatistics = new FieldStatistics();
            _fieldstatistics.Owner = FrameworkApplication.Current.MainWindow;
            _fieldstatistics.Closed += (o, e) => { _fieldstatistics = null; };
            _fieldstatistics.Show();
            //uncomment for modal
            //_fieldstatistics.ShowDialog();
        }

        protected override void OnUpdate()
        {
            QueuedTask.Run(() =>
            {
                // 预设可见
                bool isEnable = true;

                var tableView = TableView.Active;
                var selectedFields = tableView.GetSelectedFields();

                // 当选中多个字段时，不可点击
                if (selectedFields.Count != 1)
                {
                    isEnable = false;
                }

                //// 当选中字段不是数字型时，不可点击
                //string fieldName = selectedFields[0];
                //FieldDescription fieldDesc = tableView.GetFields().FirstOrDefault(x => x.Name == fieldName);
                //FieldType fieldType = fieldDesc.Type;

                //if (fieldType != FieldType.SmallInteger && fieldType != FieldType.Integer && fieldType != FieldType.Single && fieldType != FieldType.Double)
                //{
                //    isEnable = false;
                //}

                this.Enabled = isEnable;
            });
        }

    }
}
