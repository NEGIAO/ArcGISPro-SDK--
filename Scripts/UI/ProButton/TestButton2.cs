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
using ArcGIS.Core.Internal.CIM;
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
using CCTool.Scripts.ToolManagers;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using static System.Windows.Forms.MonthCalendar;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using Aspose.Cells.Drawing;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using NPOI.XWPF.UserModel;
using Aspose.Cells;
using Aspose.Words;
using CCTool.Scripts.CusTool2;
using CCTool.Scripts.ToolManagers.Managers;
using Document = Aspose.Words.Document;
using CCTool.Scripts.ToolManagers.Extensions;

namespace CCTool.Scripts.UI.ProButton
{
    internal class TestButton2 : Button
    {

        // 定义一个进度框
        private ProcessWindow processwindow = null;

        protected override async void OnClick()
        {

            // 获取参数
            string def_folder = Project.Current.HomeFolderPath;     // 工程默认文件夹位置
            string def_gdb = Project.Current.DefaultGeodatabasePath;    // 工程默认数据库

            try
            {
                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, "进度");
                pw.AddMessageTitle("XXX");


                //ToTXT(path, path1);
                //ToWord(path1, path2);

                string lyName = "用来叠加的征地批文库";
                string result = "";
                string v = "01";

                int a = v.ToInt();

                pw.AddMessageMiddle(100, $"{a}", Brushes.Blue);
                await QueuedTask.Run(() =>
                {

                });


                pw.AddMessageMiddle(100, "工具执行结束", Brushes.Blue);
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }


        
        private void ToTXT(string path, string path1)
        {
            List<string> files = DirTool.GetAllFiles(path, ".cs");
            // 转txt
            foreach (var file in files)
            {
                string fileName = Path.GetFileName(file);

                string name = fileName.Replace(".cs", ".txt");

                string targetPath = $@"{path1}\{name}";

                File.Copy(file, targetPath, true);
            }
        }

        private void ToWord(string path1, string path2)
        {
            // 转word
            List<string> files2 = DirTool.GetAllFiles(path1);
            foreach (var file in files2)
            {
                string fileName = Path.GetFileName(file);

                string name = fileName.Replace(".txt", ".docx");

                string targetPath = $@"{path2}\{name}";


                // 读取 txt 文件内容
                string content = File.ReadAllText(file);

                // 创建 Aspose.Words 文档

                // 打开Word
                string wordPath = @"C:\Users\Administrator\Desktop\cc.docx";
                Document doc = WordTool.OpenDocument(wordPath);

                DocumentBuilder builder = new DocumentBuilder(doc);

                // 写入内容到 docx
                builder.Writeln(content);

                // 保存为 docx 文件
                doc.Save(targetPath);
            }
        }

        // 查找字符在字符串中出现的所有位置
        public static List<int> GetIndexsOfString(string str, string substr)
        {
            List<int> foundItems = new List<int>();
            int startPos = 0;
            int foundPos = -1;
            int count = 0;

            do
            {
                foundPos = str.IndexOf(substr, startPos);
                if (foundPos > -1)
                {
                    startPos = foundPos + 1;
                    count++;
                    foundItems.Add(foundPos);
                }
            } while (foundPos > -1 && startPos < str.Length);


            return foundItems;
        }



    }
}

