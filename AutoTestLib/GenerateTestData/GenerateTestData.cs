using System;
using System.Text;

namespace AutoTestLib.GenerateTestData
{
    public class GenerateTestData
    {
        static Random _random = new Random();

        public static Boolean getRandomBoolean()
        {
            int random_0_or_1 = _random.Next(0, 2);
            return random_0_or_1 > 0 ? true : false;
        }

        public static int getRandomNumber(int min, int max)
        {
            return _random.Next(min, max);
        }

        public static char getRandomChar()
        {
            int num = _random.Next(0, 26); // Zero to 25
            char let = (char)('A' + num);
            return let;
        }

        public static String getRandomString(int length)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < length; i++)
            {
                sb.Append(getRandomChar());
            }
            return sb.ToString();
        }

        public static String getRandomStringIn(String[] array)
        {
            return array[getRandomNumber(0, array.Length - 1)];
        }

        public static String getDate(String format, int dateDiff)
        {
            if (format == null)
            {
                format = "MM/dd/yyyy";
            }
            else if (format.Equals("AUS") || format.Equals("UK"))
            {
                format = "dd/MM/yyyy";
            }
            else if (format.Equals("ISO"))
            {
                format = "yyyy-MM-dd";
            }

            DateTime today = DateTime.Today;
            return today.AddDays(dateDiff).ToString(format);
        }

        public static String Today(String format)
        {
            return getDate(format, 0);
        }

        public static String Tomorrow(String format)
        {
            return getDate(format, 1);
        }

        public static String Yesterday(String format)
        {
            return getDate(format, -1);
        }
    }
}
