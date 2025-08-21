using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Data.Raster;
using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Catalog;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.DataPross.GDB;
using CCTool.Scripts.Manager;
using CCTool.Scripts.MiniTool.GetInfo;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CCTool.Scripts.GDBMenu
{
    /// <summary>
    /// Interaction logic for ClearGDBAll.xaml
    /// </summary>
    public partial class ClearGDBAll : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public ClearGDBAll()
        {
            InitializeComponent();
            // 刷新表格内容
            RefrashItem();
        }

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "完全清除GDB数据库";

        // 刷新表格内容
        public async void RefrashItem()
        {
            try
            {
                // 定义一个空包
                List<ItemAtt> itemAttList = new List<ItemAtt>();
                // 获取当前选择的gdb数据库
                GDBProjectItem gdbItem = Project.Current.SelectedItems.OfType<GDBProjectItem>().FirstOrDefault();
                await QueuedTask.Run(() =>
                {
                    // 获取gdb数据库下的所有数据
                    using (Geodatabase gdb = new Geodatabase(new FileGeodatabaseConnectionPath(new Uri(gdbItem.Path))))
                    {
                        // 要素数据集
                        var dbDefinitions = gdb.GetDefinitions<FeatureDatasetDefinition>();
                        itemAttList.Add(new ItemAtt() { GDBItem = "要素数据集", Number = dbDefinitions.Count.ToString() });
                        // 要素类
                        var fcDefinitions = gdb.GetDefinitions<FeatureClassDefinition>();
                        itemAttList.Add(new ItemAtt() { GDBItem = "要素类", Number = fcDefinitions.Count.ToString() });
                        // 独立表
                        var tableDefinitions = gdb.GetDefinitions<TableDefinition>();
                        itemAttList.Add(new ItemAtt() { GDBItem = "独立表", Number = tableDefinitions.Count.ToString() });
                        // 栅格
                        var rasterDefinitions = gdb.GetDefinitions<RasterDatasetDefinition>();
                        itemAttList.Add(new ItemAtt() { GDBItem = "栅格", Number = rasterDefinitions.Count.ToString() });
                    };
                });

                // 绑定
                dg.ItemsSource = itemAttList;
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                // 定义一个空包
                List<ItemAtt> itemAttList = new List<ItemAtt>();
                // 获取当前选择的gdb数据库
                GDBProjectItem gdb = Project.Current.SelectedItems.OfType<GDBProjectItem>().FirstOrDefault();
                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart( $"获取所有数据");

                    // 要素数据集
                    List<string> dbList = gdb.Path.GetDataBasePath();
                    // 要素类
                    List<string> originFCList = gdb.Path.GetFeatureClassPathFromGDB();
                    // 独立表
                    List<string> tbList = gdb.Path.GetStandaloneTablePathFromGDB();
                    // 栅格
                    List<string> rasterList = gdb.Path.GetRasterPath();

                    // 如果要素类在要素数据集下，就剔除
                    List<string> fcList =new List<string>();
                    foreach (string fc in originFCList)
                    {
                        string fcName = fc[(fc.LastIndexOf(".gdb") + 5)..];
                        if (fcName.IndexOf(@"\") < 0)
                        {
                            fcList.Add(fc);
                        }
                    }

                    // 删除
                    foreach (var db in dbList)
                    {
                        string dbName = db[(db.LastIndexOf(".gdb") + 5)..];
                        pw.AddMessageMiddle(5, $"删除要素数据集_{dbName}");
                        Arcpy.Delect(db);
                    }
                    foreach (var fc in fcList)
                    {
                        string fcName = fc[(fc.LastIndexOf(".gdb") + 5)..];
                        pw.AddMessageMiddle(5, $"删除要素类_{fcName}");
                        Arcpy.Delect(fc);
                    }
                    foreach (var tb in tbList)
                    {
                        string tbName = tb[(tb.LastIndexOf(".gdb") + 5)..];
                        pw.AddMessageMiddle(5, $"删除独立表_{tbName}");
                        Arcpy.Delect(tb);
                    }
                    foreach (var raster in rasterList)
                    {
                        string rasterName = raster[(raster.LastIndexOf(".gdb") + 5)..];
                        pw.AddMessageMiddle(5, $"删除栅格_{rasterName}");
                        Arcpy.Delect(raster);
                    }

                });

                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }

        private void btn_cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
            return;
        }
    }

    public class ItemAtt
    {
        public string GDBItem { get; set; }
        public string Number { get; set; }
    }
}
