using ArcGIS.Core.Data;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers;
using CCTool.Scripts.ToolManagers.Extensions;
using CCTool.Scripts.UI.ProWindow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CCTool.Scripts.TableMenu
{
    /// <summary>
    /// Interaction logic for FieldAttribute.xaml
    /// </summary>
    public partial class FieldAttribute : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public FieldAttribute()
        {
            InitializeComponent();

            // 更新字段属性
            UpdataField();
        }

        // 统计字段
        private async void UpdataField()
        {
            try
            {
                // 获取字段属性
                Field field = await QueuedTask.Run(() =>
                {
                    return GisTool.GetSelectField();
                });

                // 类型转为文本型
                string fieldType = field.FieldType switch
                {
                    FieldType.String => "文本型",
                    FieldType.Integer => "长整型",
                    FieldType.SmallInteger => "短整型",
                    FieldType.Single => "单精度",
                    FieldType.Double => "双精度",
                    FieldType.OID => "OID",
                    FieldType.Geometry => "Geometry",
                    FieldType.Blob => "Blob",
                    FieldType.Date => "时间型",
                    _ => "未知",
                };

                // 写入文本框
                text_name.Text = field.Name;
                text_aliasName.Text = field.AliasName;
                text_type.Text = fieldType;
                text_length.Text = fieldType == "文本型"? field.Length.ToString():"";
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }

        private void btn_help_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/147463887";
            UITool.Link2Web(url);
        }

    }
}
