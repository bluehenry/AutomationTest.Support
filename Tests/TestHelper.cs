using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTest.AU
{
    public class TestHelper
    {
        public static String getSiteRootUrl()
        {
            // change to your installed location
            return "http:\\localhost:80";
        }
                
        public static String getDataPath()
        {
            // change to yours
            return @"C:\Workspace\CSharp\AU\Development\UnitTest.AU\";
        }

        public static String getTempPath()
        {
            return @"C:\temp";
        }
    }
}
