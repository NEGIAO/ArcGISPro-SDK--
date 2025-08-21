using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Framework.Utilities;
using ArcGIS.Desktop.Mapping;
using Aspose.Cells;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.ToolManagers.Managers;
using MathNet.Numerics;
using MathNet.Numerics.LinearAlgebra.Factorization;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
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
using Range = Aspose.Cells.Range;

namespace CCTool.Scripts.CusTool
{
    /// <summary>
    /// Interaction logic for ExcelBoundaryTable.xaml
    /// </summary>
    public partial class ExcelBoundaryTable : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        // 工具设置标签
        readonly string toolSet = "ExcelBoundaryTable";
        public ExcelBoundaryTable()
        {
            InitializeComponent();

            // 初始化参数选项
            textFolderPath.Text = BaseTool.ReadValueFromReg(toolSet, "excel_folder");
        }

        // 更新默认字段
        public async void UpdataField()
        {
            string ly = combox_fc.ComboxText();

            // 初始化参数选项
            await UITool.InitLayerFieldToComboxPlus(combox_zddm, ly, "ZDDM", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_qlr, ly, "QSDW", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_zl, ly, ["坐落", "ZL"], "string");
            await UITool.InitLayerFieldToComboxPlus(combox_fl, ly, "法人", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_sfz, ly, "法人SHZ", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_tfh, ly, "TFH", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_zdmj, ly, "ZDMJ", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_zzjgdm, ly, "组织JGDM", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_bdcdyh, ly, "BDCDYH", "string");

            await UITool.InitLayerFieldToComboxPlus(combox_txdz, ly, "TXDZ", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_frdh, ly, "法人SJH", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_bdch, ly, "原BDCZSH", "string");

            await UITool.InitLayerFieldToComboxPlus(combox_bz, ly, "ZDSZB", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_dz, ly, "ZDSZD", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_nz, ly, "ZDSZN", "string");
            await UITool.InitLayerFieldToComboxPlus(combox_xz, ly, "ZDSZX", "string");

            await UITool.InitLayerFieldToComboxPlus(combox_nyd, ly, "NYDMJ", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_gd, ly, "GDMJ", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_ld, ly, "LDMJ", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_cd, ly, "CDMJ", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_qt, ly, "QTYDMJ", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_jsyd, ly, "JSYDMJ", "float");
            await UITool.InitLayerFieldToComboxPlus(combox_wlyd, ly, "WLYDMJ", "float");
        }


        // 定义一个进度框
        private ProcessWindow processwindow = null;
        string tool_name = "权属调查表(伊)";

        private void combox_fc_DropDown(object sender, EventArgs e)
        {
            UITool.AddFeatureLayersToComboxPlus(combox_fc);
        }

        private void openFolderButton_Click(object sender, RoutedEventArgs e)
        {
            textFolderPath.Text = UITool.OpenDialogFolder();
        }

        // 获取字段值，如无则为""
        private string GetFieldValue(Feature feature, string fieldName)
        {
            string result = "";

            if (fieldName is null || fieldName == "" )
            {
                return result;
            }
            else
            {
                result = feature[fieldName].ToString();
                return result;
            }
        }


        private async void btn_go_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 参数获取
                string in_fc = combox_fc.ComboxText();
                string excel_folder = textFolderPath.Text;

                string field_zddm = combox_zddm.ComboxText();
                string field_qlr = combox_qlr.ComboxText();
                string field_zl = combox_zl.ComboxText();
                string field_fl = combox_fl.ComboxText();
                string field_sfz = combox_sfz.ComboxText();
                string field_tfh = combox_tfh.ComboxText();
                string field_zdmj = combox_zdmj.ComboxText();
                string field_zzjgdm = combox_zzjgdm.ComboxText();
                string field_bdcdyh = combox_bdcdyh.ComboxText();

                string field_txdz = combox_txdz.ComboxText();
                string field_frdh = combox_frdh.ComboxText();
                string field_bdch = combox_bdch.ComboxText();

                string field_bz = combox_bz.ComboxText();
                string field_dz = combox_dz.ComboxText();
                string field_nz = combox_nz.ComboxText();
                string field_xz = combox_xz.ComboxText();

                string field_nyd = combox_nyd.ComboxText();
                string field_gd = combox_gd.ComboxText();
                string field_ld = combox_ld.ComboxText();
                string field_cd = combox_cd.ComboxText();
                string field_qt = combox_qt.ComboxText();
                string field_jsyd = combox_jsyd.ComboxText();
                string field_wlyd = combox_wlyd.ComboxText();

                // 判断参数是否选择完全
                if (in_fc == "" || excel_folder == "" || field_zddm == "")
                {
                    MessageBox.Show("有必选参数为空！！！");
                    return;
                }

                // 写入本地
                BaseTool.WriteValueToReg(toolSet, "excel_folder", excel_folder);

                // 打开进度框
                ProcessWindow pw = UITool.OpenProcessWindow(processwindow, tool_name);
                pw.AddMessageTitle(tool_name);

                Close();

                await QueuedTask.Run(() =>
                {
                    pw.AddMessageStart("获取目标FeatureLayer");
                    // 获取目标FeatureLayer
                    FeatureLayer featurelayer = in_fc.TargetFeatureLayer();
                    // 确保要素类的几何类型是多边形
                    if (featurelayer.ShapeType != esriGeometryType.esriGeometryPolygon)
                    {
                        // 如果不是多边形类型，则输出错误信息并退出函数
                        MessageBox.Show("该要素类不是多边形类型。");
                        return;
                    }

                    // 遍历面要素类中的所有要素
                    RowCursor cursor = featurelayer.TargetSelectCursor();
                    while (cursor.MoveNext())
                    {
                        using var feature = cursor.Current as Feature;
                        // 获取参数
                        string oidField = in_fc.TargetIDFieldName();
                        string oid = feature[oidField].ToString();

                        string zddm = GetFieldValue(feature, field_zddm);

                        string qlr = GetFieldValue(feature, field_qlr); 
                        string zl = GetFieldValue(feature, field_zl); 
                        string fl = GetFieldValue(feature, field_fl); 
                        string sfz = GetFieldValue(feature, field_sfz); 
                        string tfh = GetFieldValue(feature, field_tfh); 
                        string zdmj = GetFieldValue(feature, field_zdmj); 
                        string zzjgdm = GetFieldValue(feature, field_zzjgdm); 
                        string bdcdyh = GetFieldValue(feature, field_bdcdyh); 

                        string txdz = GetFieldValue(feature, field_txdz); 
                        string frdh = GetFieldValue(feature, field_frdh); 
                        string bdch = GetFieldValue(feature, field_bdch); 

                        string bz = GetFieldValue(feature, field_bz); 
                        string dz = GetFieldValue(feature, field_dz); 
                        string nz = GetFieldValue(feature, field_nz); 
                        string xz = GetFieldValue(feature, field_xz); 

                        string nyd = GetFieldValue(feature, field_nyd); 
                        string gd = GetFieldValue(feature, field_gd); 
                        string ld = GetFieldValue(feature, field_ld); 
                        string cd = GetFieldValue(feature, field_cd); 
                        string qt = GetFieldValue(feature, field_qt); 
                        string jsyd = GetFieldValue(feature, field_jsyd); 
                        string wlyd = GetFieldValue(feature, field_wlyd); 

                        pw.AddMessageMiddle(20, $"处理要素：{oid} - 宗地码：{zddm}");

                        // 复制界址点Excel表
                        string excelPath = excel_folder + @$"\权籍调查表_{zddm}.xls";
                        DirTool.CopyResourceFile(@"CCTool.Data.Excel.权籍调查表模板.xls", excelPath);

                        // 打开工作薄
                        Workbook wb = ExcelTool.OpenWorkbook(excelPath);

                        // 处理sheet_封面
                        Worksheet ws_fm = wb.Worksheets["封面"];
                        Cells cells_fm = ws_fm.Cells;
                        cells_fm["D19"].Value = $"{zddm}";

                        // 处理sheet_基本表
                        Worksheet ws_jbb = wb.Worksheets["基本表"];
                        Cells cells_jbb = ws_jbb.Cells;
                        cells_jbb["D3"].Value = $"{qlr}";
                        cells_jbb["H6"].Value = $"{zzjgdm}";
                        cells_jbb["H7"].Value = $"{txdz}";

                        cells_jbb["C10"].Value = $"{zl}";
                        cells_jbb["H11"].Value = $"{frdh}";

                        cells_jbb["C11"].Value = $"{fl}";
                        cells_jbb["E12"].Value = $"{sfz}";
                        cells_jbb["G16"].Value = $"{zddm}W00000000";
                        cells_jbb["C17"].Value = $"{zddm}";
                        cells_jbb["G17"].Value = $"{zddm}";
                        cells_jbb["E19"].Value = $"{tfh}";
                        cells_jbb["C20"].Value = $"北：{bz}";
                        cells_jbb["C21"].Value = $"东：{dz}";
                        cells_jbb["C22"].Value = $"南：{nz}";
                        cells_jbb["C23"].Value = $"西：{xz}";
                        cells_jbb["E27"].Value = $"{zdmj}";

                        // 处理sheet_宗地分类面积调查表
                        Worksheet ws_mj = wb.Worksheets["宗地分类面积调查表"];
                        Cells cells_mj = ws_mj.Cells;
                        cells_mj["E3"].Value = $"{qlr}";
                        cells_mj["E4"].Value = $"{zddm}";
                        cells_mj["E5"].Value = $"{zddm}W00000000";
                        cells_mj["F6"].Value = $"{nyd}";
                        cells_mj["F7"].Value = $"{gd}";
                        cells_mj["F8"].Value = $"{ld}";
                        cells_mj["F9"].Value = $"{cd}";
                        cells_mj["F10"].Value = $"{qt}";
                        cells_mj["F11"].Value = $"{jsyd}";
                        cells_mj["F12"].Value = $"{wlyd}";

                        // 处理sheet_申请表1
                        Worksheet ws_s1 = wb.Worksheets["申请表1"];
                        Cells cells_s1 = ws_s1.Cells;
                        cells_s1["C10"].Value = $"{qlr}";
                        cells_s1["G11"].Value = $"{zzjgdm}";
                        cells_s1["C12"].Value = $"{zl}";
                        cells_s1["C13"].Value = $"{fl}";
                        cells_s1["G13"].Value = $"{frdh}";
                        cells_s1["D16"].Value = $"{zl}";
                        cells_s1["D17"].Value = $"{bdcdyh}";
                        cells_s1["D18"].Value = $"宗地面积：{zdmj}";
                        cells_s1["D19"].Value = $"{bdch}";

                        // 保存
                        wb.Save(excelPath);
                        wb.Dispose();

                        //// 导出PDF
                        //ExcelTool.ImportToPDF(excelPath, excel_folder + @$"\权籍调查表_{zddm}.pdf");
                    }
                    pw.AddMessageEnd();
                });
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }
        }


        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/139114694";
            UITool.Link2Web(url);
        }

        private void combox_zddm_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_zddm);
        }

        private void combox_qlr_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_qlr);
        }

        private void combox_zl_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_zl);
        }

        private void combox_fl_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_fl);
        }

        private void combox_sfz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_sfz);
        }

        private void combox_tfh_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_tfh);
        }

        private void combox_zdmj_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_zdmj);
        }

        private void combox_zzjgdm_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_zzjgdm);
        }

        private void combox_bdcdyh_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_bdcdyh);
        }

        private void combox_bz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_bz);
        }

        private void combox_dz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_dz);
        }

        private void combox_nz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_nz);
        }

        private void combox_xz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_xz);
        }

        private void combox_nyd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_nyd);
        }

        private void combox_gd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_gd);
        }

        private void combox_ld_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_ld);
        }

        private void combox_cd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_cd);
        }

        private void combox_qt_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_qt);
        }

        private void combox_jsyd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_jsyd);
        }

        private void combox_wlyd_DropDown(object sender, EventArgs e)
        {
            UITool.AddFloatFieldsToComboxPlus(combox_fc.ComboxText(), combox_wlyd);
        }

        private void combox_bdch_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_bdch);
        }

        private void combox_txdz_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_txdz);
        }

        private void combox_frdh_DropDown(object sender, EventArgs e)
        {
            UITool.AddTextFieldsToComboxPlus(combox_fc.ComboxText(), combox_frdh);
        }

        private void combox_fc_DropClose(object sender, EventArgs e)
        {
            // 更新默认字段
            UpdataField();
        }
    }
}
