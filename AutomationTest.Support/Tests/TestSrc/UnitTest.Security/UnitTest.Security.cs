using Microsoft.VisualStudio.TestTools.UnitTesting;
using AutomationUtils.Security;


namespace UnitTest.AU.TestSrc.UnitTest.Security
{
    [TestClass]
    public class UnitTest_Security
    {
        [TestMethod]
        public void TestMethod_EncrytpDecrypt()
        {

            string str = "I'm a password.";
            string key = "I'm a key!";

            string encryptedStr = "";
            string decryptedStr = "";

            encryptedStr = EncrytpDecrypt.Encrypt(str, key);

            decryptedStr = EncrytpDecrypt.Decrypt(encryptedStr, key);

            Assert.AreEqual(str, decryptedStr);

        }

    }
}
