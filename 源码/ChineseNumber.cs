using System;
using System.Globalization;

namespace CaoXing
{
    /// <summary>中文数字与阿拉伯数字互转，用于“模块一”“方面五”等标签。</summary>
    public static class Cn
    {
        private static readonly string[] Digits = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };

        public static string Number(int n)
        {
            if (n <= 0) return n.ToString(CultureInfo.InvariantCulture);
            if (n < 10) return Digits[n];
            if (n < 20) return "十" + (n % 10 == 0 ? "" : Digits[n % 10]);
            if (n < 100) return Digits[n / 10] + "十" + (n % 10 == 0 ? "" : Digits[n % 10]);
            return n.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>解析“一”“十一”“二十一”“1”等写法，失败返回 0。</summary>
        public static int Parse(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            text = text.Trim();
            int arabic;
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out arabic))
                return arabic;
            if (text == "十") return 10;
            int ten = text.IndexOf('十');
            if (ten >= 0)
            {
                string left = text.Substring(0, ten);
                string right = text.Substring(ten + 1);
                int tens = left.Length == 0 ? 1 : DigitValue(left);
                int ones = right.Length == 0 ? 0 : DigitValue(right);
                if (tens > 0 && ones >= 0) return tens * 10 + ones;
                return 0;
            }
            return DigitValue(text);
        }

        private static int DigitValue(string s)
        {
            if (s.Length != 1) return -1;
            for (int i = 0; i < Digits.Length; i++)
                if (Digits[i] == s) return i;
            return -1;
        }
    }
}
