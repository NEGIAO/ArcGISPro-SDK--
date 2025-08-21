using ArcGIS.Core.Geometry;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Managers
{
    public class BaseTool
    {
        // 把字符串写入注册表
        public static void WriteValueToReg(string Path, string keyName, object keyValue)
        {
            string keyString = keyValue.ToString();

            // 定义注册表路径
            string registryPath = $@"Software\ArcGISProData\{Path}";
            // 检查注册表中是否存在指定路径
            RegistryKey key = Registry.CurrentUser.OpenSubKey(registryPath, true);
            // 如果路径不存在，则创建路径
            key ??= Registry.CurrentUser.CreateSubKey(registryPath, true);
            // 设置或创建字符串值
            key.SetValue(keyName, keyString, RegistryValueKind.String);
            // 关闭注册表键
            key.Close();
        }

        // 从注册表中读取字符串
        public static string ReadValueFromReg(string Path, string keyName, string defValue = "")
        {
            string result = defValue;
            // 定义注册表路径
            string registryPath = $@"Software\ArcGISProData\{Path}";
            // 读取值
            RegistryKey key = Registry.CurrentUser.OpenSubKey(registryPath, true);
            if (key != null)
            {
                result = key.GetValue(keyName)?.ToString();
                // 关闭注册表键
                key.Close();
            }

            return result;
        }

        // 计算点2点距离
        public static double CalculateDistance(MapPoint pt1, MapPoint pt2)
        {
            return Math.Sqrt(Math.Pow(pt1.X - pt2.X, 2) + Math.Pow(pt1.Y - pt2.Y, 2));
        }


        // 计算点2相对于点1的角度     从正东方向为0度，-180~180
        public static double CalculateAngleFromEast(List<double> point1, List<double> point2)
        {
            double deltaX = point2[0] - point1[0];
            double deltaY = point2[1] - point1[1];
            double radians = Math.Atan2(deltaX, deltaY);
            double angle = radians * (180 / Math.PI);
            return angle;

        }

        // 计算点2相对于点1的角度     从正北方向为0度，-180~180
        public static double CalculateAngleFromNorth(List<double> point1, List<double> point2)
        {
            double result = 0;

            double deltaX = point2[0] - point1[0];
            double deltaY = point2[1] - point1[1];
            double radians = Math.Atan2(deltaX, deltaY);
            double angle = radians * (180 / Math.PI);
            // 此时的角度取值范围为-180~180, 正北方向0度
            // 调整至0-360度
            if (angle < 0)
            {
                result = angle + 360;
            }
            else
            {
                result = angle;
            }
            return result;
        }

        // 阿拉伯数字转中文数字
        public static string NumConverToChinese(string input)
        {
            // 正则表达式匹配数字
            string pattern = @"\d+";

            // 使用正则表达式查找字符串中的阿拉伯数字部分，并替换为中文数字
            string output = Regex.Replace(input, pattern, match => NumberToChinese(match.Value));

            return output;

        }


        // 中文数字转阿拉伯数字
        public static string ChineseConverToNum(string input)
        {
            // 正则表达式匹配中文数字字符（数字和单位）
            string pattern = "[零一二三四五六七八九十百千万亿]+";

            // 使用正则替换，将中文数字替换为阿拉伯数字
            string output = Regex.Replace(input, pattern, new MatchEvaluator(ReplaceChineseNumber));

            return output;

        }

        // 此方法将匹配到的中文数字字符串转换成阿拉伯数字字符串
        static string ReplaceChineseNumber(Match m)
        {
            long num = ChineseToNumber(m.Value);
            return num.ToString();
        }

        // 中文数字转阿拉伯数字
        public static long ChineseToNumber(string chineseNumber)
        {
            // 定义中文数字到数值的映射
            Dictionary<char, int> digitMap = new Dictionary<char, int>()
        {
            {'零', 0},
            {'一', 1},
            {'二', 2},
            {'三', 3},
            {'四', 4},
            {'五', 5},
            {'六', 6},
            {'七', 7},
            {'八', 8},
            {'九', 9}
        };

            // 定义中文单位到数值的映射
            Dictionary<char, long> unitMap = new Dictionary<char, long>()
        {
            {'十', 10},
            {'百', 100},
            {'千', 1000},
            {'万', 10000},
            {'亿', 100000000}
        };

            long result = 0;    // 最终结果
            long section = 0;   // 用于处理“万”或“亿”等较大单位前的一段
            long number = 0;    // 当前读到的数字

            foreach (char c in chineseNumber)
            {
                if (digitMap.ContainsKey(c))
                {
                    // 读取到数字，保存对应数值
                    number = digitMap[c];
                }
                else if (unitMap.ContainsKey(c))
                {
                    long unit = unitMap[c];
                    // 如果遇到万、亿这样的单位，则先将前面的部分与该单位相乘后加入总结果
                    if (unit == 10000 || unit == 100000000)
                    {
                        section = (section + number) * unit;
                        result += section;
                        section = 0;
                    }
                    else
                    {
                        // 如果当前没有前面的数字，则默认数字为1（例如“十”代表10而非0）
                        section += (number == 0 ? 1 : number) * unit;
                    }
                    number = 0;
                }
            }
            // 加上最后可能存在的数字
            return result + section + number;
        }

        // 阿拉伯数字转中文数字
        public static string NumberToChinese(string number)
        {
            string[] numArray = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
            string[] unitArray = { "", "十", "百", "千" };
            string[] sectionArray = { "", "万", "亿", "万亿" };

            int length = number.Length;
            int sectionCount = (length + 3) / 4; // 计算需要的节数
            string result = "";
            bool needZero = false; // 标记是否需要添加“零”

            for (int i = 0; i < sectionCount; i++)
            {
                int sectionIndex = sectionCount - i - 1;
                int sectionLength = length - i * 4 > 4 ? 4 : length - i * 4;
                string section = number.Substring(length - i * 4 - sectionLength, sectionLength);
                int sectionNum = int.Parse(section);

                if (sectionNum == 0)
                {
                    needZero = true;
                    continue;
                }

                string sectionResult = "";
                bool sectionZero = false; // 标记节内是否需要添加“零”

                for (int j = 0; j < section.Length; j++)
                {
                    int digitIndex = section.Length - j - 1;
                    int digit = int.Parse(section[j].ToString());

                    if (digit == 0)
                    {
                        sectionZero = true;
                    }
                    else
                    {
                        if (sectionZero)
                        {
                            sectionResult += numArray[0];
                            sectionZero = false;
                        }
                        sectionResult += numArray[digit] + unitArray[digitIndex];
                    }
                }

                if (needZero)
                {
                    result += numArray[0];
                    needZero = false;
                }

                result += sectionResult + sectionArray[sectionIndex];
            }

            // 处理“十”开头的情况，例如“10”转换为“十”
            if (result.StartsWith("一十"))
            {
                result = result.Substring(1);
            }

            return result;
        }
    }
}
