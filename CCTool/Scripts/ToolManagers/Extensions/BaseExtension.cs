using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Extensions
{
    public static class BaseExtension
    {
        // 提取特定文字【model包括：中文、英文、数字、特殊符号】
        public static string GetWord(this string txt_in, string model = "中文")
        {
            string chinesePattern = "[\u4e00-\u9fa5]"; // 匹配中文字符的正则表达式
            string englishPattern = "[a-zA-Z]"; // 匹配英文字符的正则表达式
            string digitPattern = @"\d"; // 匹配数字的正则表达式
            string specialCharPattern = @"[^a-zA-Z0-9\u4e00-\u9fa5\s]"; // 匹配特殊符号的正则表达式

            string decimalPattern = @"\d+\.?\d*";  // 匹配整数和小数（如3.25、6.0）

            string txt = "";

            if (model == "中文")
            {
                Regex chineseRegex = new Regex(chinesePattern);
                txt = ExtractMatches(txt_in, chineseRegex);
            }
            else if (model == "英文")
            {
                Regex englishRegex = new Regex(englishPattern);
                txt = ExtractMatches(txt_in, englishRegex);
            }
            else if (model == "数字")
            {
                Regex digitRegex = new Regex(digitPattern);
                txt = ExtractMatches(txt_in, digitRegex);
            }
            else if (model == "小数")
            {
                Regex digitRegex = new Regex(decimalPattern);
                txt = ExtractMatches(txt_in, digitRegex);
            }
            else if (model == "特殊符号")
            {
                Regex specialCharRegex = new Regex(specialCharPattern);
                txt = ExtractMatches(txt_in, specialCharRegex);
            }
            return txt;
        }
        // 正则匹配
        public static string ExtractMatches(string input, Regex regex)
        {
            string result = "";
            MatchCollection matches = regex.Matches(input);

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                result += match.Value;
            }
            return result;
        }

        // 填充小数位数，不够就用0代替
        public static string RoundWithFill(this double value, int digit)
        {
            // 要填充的0的位数
            int lengZero = 0;

            // 先Round一下
            double va = Math.Round(value, digit);
            // 转成字符串处理
            string txt_value = va.ToString();
            // 分析一下，如果只有整数位，即没有小数点的情况
            if (!txt_value.Contains("."))
            {
                lengZero = digit;
                txt_value += ".";
            }
            // 有小数位数的时候
            else
            {
                lengZero = digit - (txt_value.Length - (txt_value.IndexOf(".") + 1));
            }

            txt_value += new string('0', lengZero);
            // 返回值
            return txt_value;
        }

        // 对重复要素进行数字标记
        public static List<string> AddNumbersToDuplicates(this List<string> stringList)
        {
            // 使用Dictionary来跟踪每个字符串的出现次数
            Dictionary<string, int> stringCount = new Dictionary<string, int>();

            // 遍历字符串列表
            for (int i = 0; i < stringList.Count; i++)
            {
                string currentString = stringList[i];

                // 检查字符串是否已经在Dictionary中存在
                if (stringCount.ContainsKey(currentString))
                {
                    // 获取该字符串的出现次数
                    int count = stringCount[currentString];

                    // 在当前字符串后添加数字
                    stringList[i] = $"{currentString}：{count + 1}";

                    // 更新Dictionary中的计数
                    stringCount[currentString] = count + 1;
                }
                else
                {
                    // 如果字符串在Dictionary中不存在，将其添加，并将计数设置为1
                    stringCount.Add(currentString, 1);
                    // 在当前字符串后添加数字
                    stringList[i] = $"{currentString}：{1}";
                }
            }
            // 去除单个要素的数字标记
            foreach (var item in stringCount)
            {
                if (item.Value == 1)
                {
                    for (int i = 0; i < stringList.Count; i++)
                    {
                        if (stringList[i] == item.Key + "：1")
                        {
                            stringList[i] = item.Key;
                        }
                    }
                }
            }

            // 返回字符串列表
            return stringList;
        }

        // 度分秒转十进制度
        public static string ToDecimal(this double cod)
        {
            double value = double.Parse(cod.ToString());
            // 计算度分秒的值
            int degree = (int)(value / 1);
            int minutes = (int)(value % 1 * 60 / 1);
            double seconds = (value % 1 * 60 - minutes) * 60;
            // 合并为字符串
            string dec = degree.ToString() + "°" + minutes.ToString() + "′" + seconds.ToString("0.0000") + "″";

            // 返回
            return dec;
        }

        // 十进制度转度分秒
        public static double ToDegree(this string dec)
        {
            // 初始化度分秒符号的位置
            int index1 = -1;
            int index2 = -1;
            int index3 = -1;
            // 定义度分秒可能的符号
            List<string> list_degree = new List<string>() { "度", "°" };
            List<string> list_minutes = new List<string>() { "分", "′", "'" };
            List<string> list_seconds = new List<string>() { "秒", "″", "\"" };
            // 找到度分秒符号的位置
            foreach (var item in list_degree)
            {
                if (dec.ToString().IndexOf(item) != -1)
                {
                    index1 = dec.ToString().IndexOf(item);
                }
            }
            foreach (var item in list_minutes)
            {
                if (dec.ToString().IndexOf(item) != -1)
                {
                    index2 = dec.ToString().IndexOf(item);
                }
            }
            foreach (var item in list_seconds)
            {
                if (dec.ToString().IndexOf(item) != -1)
                {
                    index3 = dec.ToString().IndexOf(item);
                }
            }
            // 计算度分秒数值
            double degree = double.Parse(dec.ToString().Substring(0, index1));
            double minutes = double.Parse(dec.ToString().Substring(index1 + 1, index2 - index1 - 1));
            double seconds = double.Parse(dec.ToString().Substring(index2 + 1, index3 - index2 - 1));
            // 计算赋值
            double deg = degree + minutes / 60 + seconds / 3600;

            // 返回
            return deg;
        }

        // 分解三级用地编码
        public static List<string> DecomposeBM(this string bm)
        {
            List<string> result = new List<string>();

            if (bm.Length >= 2)
            {
                result.Add(bm[..2]);
            }
            if (bm.Length >= 4)
            {
                result.Add(bm[..4]);
            }
            if (bm.Length >= 6)
            {
                result.Add(bm[..6]);
            }
            return result;
        }

        // 给字典增加子顶，如果已有，则累加
        public static void Accumulation(this Dictionary<string, double> dict, string key, double value)
        {
            // 复制一个
            if (!dict.ContainsKey(key))
            {
                dict.Add(key, value);
            }
            else
            {
                dict[key] += value;
            }
        }

        // 判断一个string是不是以数字开头
        public static bool IsNumeric(this string st)
        {
            string str = st[0].ToString();
            if (str == null || str.Length == 0)    //验证这个参数是否为空
                return false;                           //是，就返回False
            ASCIIEncoding ascii = new ASCIIEncoding();//new ASCIIEncoding 的实例
            byte[] bytestr = ascii.GetBytes(str);         //把string类型的参数保存到数组里

            foreach (byte c in bytestr)                   //遍历这个数组里的内容
            {
                if (c < 48 || c > 57)                          //判断是否为数字
                {
                    return false;                              //不是，就返回False
                }
            }
            return true;                                        //是，就返回True
        }

        // string转为Int
        public static int ToInt(this string st, int defaultPar = 0)
        {
            double db = st.ToDouble(defaultPar);

            return (int)db;
        }

        // string转为Long
        public static long ToLong(this string st, long defaultPar = 0)
        {
            double db = st.ToDouble(defaultPar);

            return (long)db;
        }

        // string转为double
        public static double ToDouble(this string st, double defaultPar = 0)
        {
            bool bo = double.TryParse(st, out double result);
            // 如果转换失败，并且有设置默认值，就按默认值
            if (!bo && defaultPar != 0)
            {
                return defaultPar;
            }
            else
            {
                return result;
            }
        }

        // string转为bool
        public static bool ToBool(this string st, string defaultPar = "")
        {
            bool result = false;

            // 如果输入值为空值，且有设置默认值时，就替换
            if (st is null || st == "")
            {
                if (defaultPar != "")
                {
                    st = defaultPar;
                }
            }
            else
            {
                // 判断
                if (st.ToLower() == "true")
                {
                    result = true;
                }
            }
            return result;
        }


        // 科学计数法文字转为double
        public static double SciToDouble(this string sci, double defaultPar = 0)
        {
            double result = 0;

            // 如果转换失败，并且有设置默认值，就按默认值
            if (defaultPar != 0)
            {
                return defaultPar;
            }
            else
            {
                return result;
            }
        }
    }
}
