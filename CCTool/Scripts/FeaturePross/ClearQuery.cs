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
using CCTool.Scripts.ToolManagers.Extensions;

namespace CCTool.Scripts.FeaturePross
{
	internal class ClearQuery : Button
	{
        protected override async void OnClick()
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    // 获取活动地图视图中选定的要素集合
                    var selectedSet = MapView.Active.Map.GetSelection();
                    // 将选定的要素集合转换为字典形式
                    var selectedList = selectedSet.ToDictionary();

                    // 收集当前选择的图层
                    foreach (var selected in selectedList)
                    {
                        // 获取图层和关联的对象 ID
                        FeatureLayer featureLayer = selected.Key as FeatureLayer;

                        // 清除属性定义
                        featureLayer.RemoveAllDefinitionQueries();

                    }

                });



            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }
	}
}
