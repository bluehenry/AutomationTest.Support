using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using AutoTestLib.VisualComparison;

namespace UnitTest.AU.TestSrc.UnitTest
{
    [TestClass]
    public class UnitTest_VisualComprasion
    {
        [TestMethod]
        public void TestMethod_ScreenCapture()
        {
            string filePath1 = @"..\..\TestData\VisualComprasion\screenshoot1.png";
            string filePath2 = @"..\..\TestData\VisualComprasion\screenshoot2.png";
            //string filePath3 = @"..\..\TestData\VisualComprasion\screenshoot3.png";

            VisualComprasion.ScreenCapture(filePath1);
            VisualComprasion.ScreenCapture(filePath2);
                        
            Assert.IsTrue(VisualComprasion.ImageCompareString(filePath1, filePath2));
        }
    }
}
