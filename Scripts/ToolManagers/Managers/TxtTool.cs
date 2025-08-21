using ArcGIS.Desktop.Core;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Managers
{
    public class TxtTool
    {
        // 获取TXT文件的编码类型
        public static System.Text.Encoding GetEncodingType(string FILE_NAME)
        {
            FileStream fs = new FileStream(FILE_NAME, FileMode.Open, FileAccess.Read);
            Encoding r = GetType(fs);
            fs.Close();
            return r;
        }

        // 通过给定的文件流，判断文件的编码类型 
        public static System.Text.Encoding GetType(FileStream fs)
        {
            byte[] Unicode = new byte[] { 0xFF, 0xFE, 0x41 };
            byte[] UnicodeBIG = new byte[] { 0xFE, 0xFF, 0x00 };
            byte[] UTF8 = new byte[] { 0xEF, 0xBB, 0xBF }; //带BOM 
            Encoding reVal = Encoding.Default;

            BinaryReader r = new BinaryReader(fs, System.Text.Encoding.Default);
            int i;
            int.TryParse(fs.Length.ToString(), out i);
            byte[] ss = r.ReadBytes(i);
            if (IsUTF8Bytes(ss) || (ss[0] == 0xEF && ss[1] == 0xBB && ss[2] == 0xBF))
            {
                reVal = Encoding.UTF8;
            }
            else if (ss[0] == 0xFE && ss[1] == 0xFF && ss[2] == 0x00)
            {
                reVal = Encoding.BigEndianUnicode;
            }
            else if (ss[0] == 0xFF && ss[1] == 0xFE && ss[2] == 0x41)
            {
                reVal = Encoding.Unicode;
            }
            r.Close();
            return reVal;

        }

        // 判断是否是不带 BOM 的 UTF8 格式 
        private static bool IsUTF8Bytes(byte[] data)
        {
            int charByteCounter = 1; //计算当前正分析的字符应还有的字节数 
            byte curByte; //当前分析的字节. 
            for (int i = 0; i < data.Length; i++)
            {
                curByte = data[i];
                if (charByteCounter == 1)
                {
                    if (curByte >= 0x80)
                    {
                        //判断当前 
                        while (((curByte <<= 1) & 0x80) != 0)
                        {
                            charByteCounter++;
                        }
                        //标记位首位若为非0 则至少以2个1开始 如:110XXXXX...........1111110X 
                        if (charByteCounter == 1 || charByteCounter > 6)
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    //若是UTF-8 此时第一位必须为1 
                    if ((curByte & 0xC0) != 0x80)
                    {
                        return false;
                    }
                    charByteCounter--;
                }
            }
            if (charByteCounter > 1)
            {
                throw new Exception("非预期的byte格式");
            }
            return true;
        }

        // 获取txt文件的文本内容
        public static string GetTXTContent(string txtFile)
        {
            // 预设文本内容
            string text = "";
            // 获取txt文件的编码方式
            Encoding encoding = TxtTool.GetEncodingType(txtFile);
            // 读取【ANSI和UTF-8】的不同+++++++（ANSI为0，UTF-8为3）
            // 我也不知道具体原理，只是找出差异点作个判断，以后再来解决这个问题------
            int encoding_index = int.Parse(encoding.Preamble.ToString().Substring(encoding.Preamble.ToString().Length - 2, 1));

            if (encoding_index == 0)        // ANSI编码的情况
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using StreamReader sr = new StreamReader(txtFile, Encoding.GetEncoding("GBK"));
                text = sr.ReadToEnd();
            }
            else if (encoding_index == 3)               // UTF8编码的情况
            {
                using StreamReader sr = new StreamReader(txtFile, Encoding.UTF8);
                text = sr.ReadToEnd();
            }
            return text;
        }

        // 从Json文本中获取属性
        public static string GetAttFormTxtJson(string txtPath, string attName)
        {
            // 读取文件内容
            string filePath = txtPath;
            string jsonContent = File.ReadAllText(filePath);

            // 解析 JSON 数据
            JObject jsonObject = JObject.Parse(jsonContent);

            // 获取 "dpi" 属性的值
            string Value = (string)jsonObject[attName];

            return Value;
        }

        // 获取特定符号包裹的内容
        public static string GetStringInside(string input, string symText = "[]")
        {
            string result = "";
            // 分解符号
            var s1 = symText[0];
            var s2 = symText[1];
            // 设置标记
            bool insideBraces = false;
            int start = 0;
            // 循环，找到起始和完结符号
            for (int i = 0; i < input.Length; i++)
            {
                if (input[i] == s1)
                {
                    insideBraces = true;
                    start = i + 1;
                }
                else if (input[i] == s2 && insideBraces)
                {
                    insideBraces = false;
                    int length = i - start;
                    string content = input.Substring(start, length);
                    string re = s1 + content + s2 + ",";
                    if (!result.Contains(content))
                    {
                        result += re;
                    }
                }
            }
            return result[..(result.Length - 1)];
        }

        // 获取子字符串在目标字符中出现的次数
        public static int StringInCount(string targetString, string subString)
        {
            int count = 0;
            // 获取第一次出现的位置
            int index = targetString.IndexOf(subString);
            while (index != -1)
            {
                count++;
                // 继续从获取到的位置后面开始取值
                index = targetString.IndexOf(subString, index + subString.Length);
            }
            // 返回出现的次数
            return count;
        }

        // 英文二十六进制转换
        public static string NumberChange(long number)
        {
            string result = "";

            while (number > 0)
            {
                number--; // 将范围从1-26调整为0-25
                long remainder = number % 26;
                result = (char)('A' + remainder) + result;
                number /= 26;
            }
            return result;
        }


        /// <summary>
        /// 转换数字
        /// </summary>
        public static long CharToNumber(char c)
        {
            switch (c)
            {
                case '一': return 1;
                case '二': return 2;
                case '三': return 3;
                case '四': return 4;
                case '五': return 5;
                case '六': return 6;
                case '七': return 7;
                case '八': return 8;
                case '九': return 9;
                case '零': return 0;
                default: return -1;
            }
        }

        /// <summary>
        /// 转换单位
        /// </summary>
        public static long CharToUnit(char c)
        {
            switch (c)
            {
                case '十': return 10;
                case '百': return 100;
                case '千': return 1000;
                case '万': return 10000;
                case '亿': return 100000000;
                default: return 1;
            }
        }
        /// <summary>
        /// 将中文数字转换阿拉伯数字
        /// </summary>
        /// <param name="cnum">汉字数字</param>
        /// <returns>长整型阿拉伯数字</returns>
        public static long ParseCnToInt(string cnum)
        {
            cnum = Regex.Replace(cnum, "\\s+", "");
            long firstUnit = 1;//一级单位                
            long secondUnit = 1;//二级单位 
            long tmpUnit = 1;//临时单位变量
            long result = 0;//结果
            for (int i = cnum.Length - 1; i > -1; --i)//从低到高位依次处理
            {
                tmpUnit = CharToUnit(cnum[i]);//取出此位对应的单位
                if (tmpUnit > firstUnit)//判断此位是数字还是单位
                {
                    firstUnit = tmpUnit;//是的话就赋值,以备下次循环使用
                    secondUnit = 1;
                    if (i == 0)//处理如果是"十","十一"这样的开头的
                    {
                        result += firstUnit * secondUnit;
                    }
                    continue;//结束本次循环
                }
                else if (tmpUnit > secondUnit)
                {
                    secondUnit = tmpUnit;
                    continue;
                }
                result += firstUnit * secondUnit * CharToNumber(cnum[i]);//如果是数字,则和单位想乘然后存到结果里
            }
            return result;
        }
    }
}
