using CCTool.Scripts.Manager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Managers
{
    public class FieldCalTool
    {
        

        // 表中的值从平方米改为公顷、平方公里或亩
        public static void ChangeUnit(string in_data, string field, int unit = 1)
        {
            // 选择修改的单位
            double cg = 10000;          // 公顷
            if (unit == 2)
            {
                cg = 1000000;        // 平方公里
            }
            else if (unit == 3)
            {
                cg = 666.66667;       // 亩
            }
            else if (unit == 4)
            {
                cg = 1;       // 平方米
            }
            // 单位换算
            Arcpy.CalculateField(in_data, field, @$"!{field}!/{cg}");
        }


        // 清除字符串字段值中的空格
        public static string ClearTextSpace(string input_table, string field_name)
        {
            string exp = $"!{field_name}!.replace(' ','')";
            Arcpy.CalculateField(input_table, field_name, exp);
            return input_table;
        }

        // 清除字符串字段值中的空值，字符串转为""
        public static string ClearTextNull(string input_table, string field_name)
        {
            string block = @"def ss(a):
                                            if a is None:
                                                return ''
                                            else:
                                                return a";

            Arcpy.CalculateField(input_table, field_name, $"ss(!{field_name}!)", block);
            return input_table;
        }

        // 清除所选字段值中的空值,多字段
        public static string ClearTextNull(string initlayer, List<string> fieldNames)
        {
            foreach (string fieldName in fieldNames)
            {
                string block = @"def ss(a):
                                            if a is None:
                                                return ''
                                            else:
                                                return a";

                Arcpy.CalculateField(initlayer, fieldName, $"ss(!{fieldName}!)", block);
            }
            return initlayer;
        }

        // 清除数字值字段值中的空值，数字型转为0
        public static string ClearMathNull(string input_table, string field_name)
        {
            string block = @"def ss(a):
                                            if a is None:
                                                return 0
                                            else:
                                                return a";

            Arcpy.CalculateField(input_table, field_name, $"ss(!{field_name}!)", block);
            return input_table;
        }

        // 将数字0转换为空值
        public static string Zero2Null(string input_table, string field_name)
        {
            string block = @"def ss(a):
                                            if a==0:
                                                return None
                                            else:
                                                return a";
            Arcpy.CalculateField(input_table, field_name, $"ss(!{field_name}!)", block);
            return input_table;
        }

        // 提取中文
        public static string GetChinese(string input_table, string input_field, string output_field)
        {
            string block = "import re\r\ndef ss(a):\r\n    va = re.findall(u'[\\u4e00-\\u9fa5]+',a)\r\n    if len(va) == 0:\r\n        result = ''\r\n    else:\r\n        result = ''\r\n        for i in range(0, len(va)):\r\n            result += va[i]\r\n    return result";

            Arcpy.CalculateField(input_table, output_field, $"ss(!{input_field}!)", block);
            return input_table;
        }

        // 提取英文
        public static string GetEnglish(string input_table, string input_field, string output_field)
        {
            string block = "import re\r\ndef ss(a):\r\n    va = re.findall(r'[a-zA-Z]', a)\r\n    if len(va) == 0:\r\n        result = ''\r\n    else:\r\n        result = ''\r\n        for i in range(0, len(va)):\r\n            result += va[i]\r\n    return result";

            Arcpy.CalculateField(input_table, output_field, $"ss(!{input_field}!)", block);
            return input_table;
        }
        // 提取数字
        public static string GetNumber(string input_table, string input_field, string output_field)
        {
            string block = "import re\r\ndef ss(a):\r\n    va = re.findall(r'\\d', a)\r\n    if len(va) == 0:\r\n        result = ''\r\n    else:\r\n        result = ''\r\n        for i in range(0, len(va)):\r\n            result += va[i]\r\n    return result";

            Arcpy.CalculateField(input_table, output_field, $"ss(!{input_field}!)", block);
            return input_table;
        }
        // 提取特殊符号
        public static string GetSymbol(string input_table, string input_field, string output_field)
        {
            string block = "import re\r\ndef ss(a):\r\n    va = re.findall(r'\\W', a)\r\n    if len(va) == 0:\r\n        result = ''\r\n    else:\r\n        result = ''\r\n        for i in range(0, len(va)):\r\n            result += va[i]\r\n    return result";

            Arcpy.CalculateField(input_table, output_field, $"ss(!{input_field}!)", block);
            return input_table;
        }
    }
}
