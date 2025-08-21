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
using CCTool.Scripts.Manager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Drawing.Imaging;
using Button = ArcGIS.Desktop.Framework.Contracts.Button;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;
using System.Windows.Forms;
using ArcGIS.Desktop.Mapping;
using Microsoft.Office.Core;
using ArcGIS.Core.Data.DDL;
using FieldDescription = ArcGIS.Core.Data.DDL.FieldDescription;
using Row = ArcGIS.Core.Data.Row;
using ArcGIS.Desktop.Editing.Attributes;
using System.Security.Cryptography;
using ArcGIS.Desktop.Editing.Templates;
using System.Windows.Input;
using System.Windows.Media.Media3D;
using System.Net;
using System.Net.Http;
using System.Windows.Media;
using System.Windows.Shapes;
using Path = System.IO.Path;
using ArcGIS.Desktop.Internal.Catalog.Wizards;
using ArcGIS.Desktop.Internal.Layouts.Utilities;
using System.Windows.Documents;
using ActiproSoftware.Windows;
using System.Windows;
using System.Runtime.InteropServices;
using ArcGIS.Desktop.Internal.Mapping.Locate;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.GeoProcessing;
using Geometry = ArcGIS.Core.Geometry.Geometry;
using Polygon = ArcGIS.Core.Geometry.Polygon;
using ArcGIS.Core.Data.Exceptions;
using Table = ArcGIS.Core.Data.Table;
using SpatialReference = ArcGIS.Core.Geometry.SpatialReference;
using System.Drawing;
using Brushes = System.Windows.Media.Brushes;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using CCTool.Scripts.ToolManagers;
using NPOI.SS.Util;
using NPOI.SS.Formula.Functions;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using MathNet.Numerics.LinearAlgebra.Factorization;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.Util;
using Org.BouncyCastle.X509;
using Org.BouncyCastle.Math.EC;
using ArcGIS.Core.Internal.CIM;
using Polyline = ArcGIS.Core.Geometry.Polyline;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using static Org.BouncyCastle.Math.Primes;
using NPOI.HPSF;
using Microsoft.VisualBasic;
using System.Threading;
using NPOI.XSSF.Streaming.Values;
using ActiproSoftware.Windows.Shapes;
using Org.BouncyCastle.Tsp;
using SharpCompress.Common;
using static NPOI.POIFS.Crypt.CryptoFunctions;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using QueryFilter = ArcGIS.Core.Data.QueryFilter;
using CCTool.Scripts.UI.ProMapTool;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;
using CCTool.Scripts.ToolManagers.Library;

namespace CCTool.Scripts.UI.ProButton
{
    internal class TestButton : Button
    {

        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "计算面积";

        protected override async void OnClick()
        {
            //Ogr.RegisterAll();// 注册所有的驱动
            //Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");
            //Gdal.SetConfigOption("SHAPE_ENCODING", "");

            // 获取参数
            string def_folder = Project.Current.HomeFolderPath;     // 工程默认文件夹位置
            string def_gdb = Project.Current.DefaultGeodatabasePath;    // 工程默认数据库

            try
            {
                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, "进度");
                pw.AddMessageTitle("创建内存GDB");

                //string excelPath = @"C:\Users\Administrator\Desktop\导出\权籍调查表_411102001001GB00001.xls";
                //string pdfPath = @"C:\Users\Administrator\Desktop\伊\权籍调查表_411102001001GB00001.pdf";


                string lyName = "用地模板";

                string excelPath = $@"C:\Users\Administrator\Desktop\FE.xlsx";
                List<string> values = new List<string>();
                var yds = GlobalData.dic_ydyh_new.ToList();
                foreach (var yd in yds)
                {
                    values.Add(yd.Key);
                    values.Add(yd.Value);
                    values.Add($"{yd.Key}{yd.Value}");
                }


                await QueuedTask.Run(() =>
                {
                    FeatureLayer featureLayer = lyName.TargetFeatureLayer();

                    // 应用符号系统
                    LayerDocument lyrFile = new LayerDocument(@"D:\【软件资料】\GIS相关\ArcGIS Pro二次开发工具\CCTool\Data\Layers\国空用地新版_2.lyrx");

                    CIMLayerDocument cimLyrDoc = lyrFile.GetCIMLayerDocument();

                    CIMUniqueValueRenderer uvr = ((CIMFeatureLayer)cimLyrDoc.LayerDefinitions[0]).Renderer as CIMUniqueValueRenderer;

                    uvr.Fields = new string[] { "JQDLBM" };

                    // 创建新的Classes集合（保留原类 + 修改后的复制类）

                    List<CIMUniqueValueClass> modifiedClasses = new List<CIMUniqueValueClass>();

                    Dictionary<string, string> dict = GlobalData.dic_ydyh_new.ToDictionary(vp => vp.Value, vp => vp.Key);


                    // 修改每个标注类别的表达式
                    foreach (CIMUniqueValueClass uvClass in uvr.Groups[0].Classes)
                    {
                        var va = uvClass.Values[0].FieldValues[0].ToString();


                        CIMUniqueValueClass moClass = uvClass.Copy();
                        moClass.Values[0].FieldValues[0] = dict[va];
                        moClass.Label = dict[va];
                        modifiedClasses.Add(moClass); // 增加更改的类

                        // 补零
                        CIMUniqueValueClass moClass3 = uvClass.Copy();
                        moClass3.Values[0].FieldValues[0] = dict[va].PadRight(6, '0');
                        moClass3.Label = dict[va].PadRight(6, '0');
                        modifiedClasses.Add(moClass3); // 增加更改的类

                        uvClass.Label = va;
                        modifiedClasses.Add(uvClass); // 保留原类

                        CIMUniqueValueClass moClass2 = uvClass.Copy();
                        moClass2.Values[0].FieldValues[0] = dict[va] + va;
                        moClass2.Label = dict[va] + va;
                        modifiedClasses.Add(moClass2); // 增加更改的类

                    }

                    uvr.Groups[0].Classes = modifiedClasses.ToArray();
                    // 应用渲染器
                    featureLayer.SetRenderer(uvr);





                    //using var cursor = featureLayer.Search();

                    //int index = 0;
                    //while (cursor.MoveNext())
                    //{
                    //    if (index>= values.Count)
                    //    { 
                    //        break;
                    //    }
                    //    Row row = cursor.Current;
                    //    row["JQDLBM"] = values[index];

                    //    index++;
                    //    row.Store();
                    //}           

                });

                pw.AddMessageEnd();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }



    }
}
