using CCTool.Scripts.Manager;
using CCTool.Scripts.ToolManagers.Managers;
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

namespace CCTool.Scripts.MiniTool.MTool
{
    /// <summary>
    /// Interaction logic for Num2Chinese.xaml
    /// </summary>
    public partial class Num2Chinese : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public Num2Chinese()
        {
            InitializeComponent();
        }

        private void txt_num_Changed(object sender, TextChangedEventArgs e)
        {
            try
            {
                // 获取阿拉伯数字
                bool result = long.TryParse(txt_num.Text, out long num);
                if (result)
                {
                    // 转中文
                    txt_chinese_simple.Text = ConverToChineseSimple(num);

                    // 转中文_繁体
                    txt_chinese_traditional.Text =ConverToChineseTraditional(num);

                    // 转罗马数字
                    txt_roman_num.Text = ConverToRoman(num);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message + ee.StackTrace);
                return;
            }

        }


        private void btn_help_Click(object sender, RoutedEventArgs e)
        {
            string url = "https://blog.csdn.net/xcc34452366/article/details/144562417";
            UITool.Link2Web(url);
        }

        // 转中文
        public static string ConverToChineseSimple(long number)
        {
            if (number > 999999999)
                return "数字太大无法转换";

            string[] chineseNumbers = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
            string[] units = { "", "十", "百", "千", "万", "十", "百", "千", "亿" };

            string result = "";
            int unitPlace = 0; // 位数

            while (number > 0)
            {
                long part = number % 10;
                if (part != 0)
                {
                    result = chineseNumbers[part] + (unitPlace == 0 ? "" : units[unitPlace]) + result;
                }
                else if (result != "" && !result.StartsWith("零"))
                {
                    result = "零" + result;
                }

                number /= 10;
                unitPlace++;
            }

            if (result == "")
            {
                result = chineseNumbers[0];
            }
            else if (result.StartsWith("零") && result.Length > 1)
            {
                result = result.Substring(1);
            }

            return result;

        }


        // 转中文_繁体
        public static string ConverToChineseTraditional(long number)
        {
            if (number > 999999999)
                return "数字太大无法转换";

            string[] chineseNumbers = { "零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖" };
            string[] units = { "", "拾", "佰", "仟", "万", "拾", "佰", "仟", "亿" };

            string result = "";
            int unitPlace = 0; // 位数

            while (number > 0)
            {
                long part = number % 10;
                if (part != 0)
                {
                    result = chineseNumbers[part] + (unitPlace == 0 ? "" : units[unitPlace]) + result;
                }
                else if (result != "" && !result.StartsWith("零"))
                {
                    result = "零" + result;
                }

                number /= 10;
                unitPlace++;
            }

            if (result == "")
            {
                result = chineseNumbers[0];
            }
            else if (result.StartsWith("零") && result.Length > 1)
            {
                result = result.Substring(1);
            }

            return result;

        }


        // 转罗马数字
        public static string ConverToRoman(long value)
        {
            string result = "";

            if (value < 1 || value > 3999)
            {
                result = "数字不在正确区间内【1-3999】";
            }
            else
            {
                int[] nums = new int[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
                string[] romans = new string[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };

                int start = 0;
                while (value > 0)
                {
                    for (int i = start, l = nums.Length; i < l; ++i)
                    {
                        if (value >= nums[i])
                        {
                            value -= nums[i];
                            result += romans[i];
                            start = i;
                            break;
                        }
                    }
                }
            }

            return result;
        }
    }
}
