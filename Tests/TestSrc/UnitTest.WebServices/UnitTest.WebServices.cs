using Microsoft.VisualStudio.TestTools.UnitTesting;
using AutoTestLib.WebServices;

namespace UnitTest.AU.TestSrc.UnitTest.WebServices
{
    /// <summary>
    /// Summary description for UnitTest
    /// </summary>
    [TestClass]
    public class UnitTest
    {
        public UnitTest()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void TestMethod_callSOAPService()
        {
            string url = @"http://www.webservicex.com/globalweather.asmx";
            string SOAPAction = @"http://www.webserviceX.NET/GetWeather";
            
            string SOAPPayloadFilePath = @"..\..\TestData\WebServices\GetWeather.xml";  
            string SOAPResult = WebServicesUtils.callSOAPService(url, SOAPAction, SOAPPayloadFilePath);

        }
    }
}
