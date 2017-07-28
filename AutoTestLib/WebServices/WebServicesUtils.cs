using System;
using System.Xml;
using System.Net;
using System.IO;

namespace AutoTestLib.WebServices
{
    public class WebServicesUtils
    {



        public static string callSOAPService(string url, string SOAPAction, string SOAPPayloadFilePath)
        {
            string SOAPResult = null;

            //Read Soap payload file

            XmlDocument xmlDoc = new XmlDocument();
            using (FileStream XMLStream = new FileStream(SOAPPayloadFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                xmlDoc.Load(XMLStream);
      

                //Create a Web request
                HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);
                webRequest.Headers.Add("SOAPAction", SOAPAction);
                webRequest.ContentType = "text/xml;charset=\"utf-8\"";
                webRequest.Accept = "text/xml";
                webRequest.Method = "POST";
                webRequest.Timeout = 300000;

                using (Stream stream = webRequest.GetRequestStream())
                {
                    //Insert Soap Envelope into WebRequest
                    xmlDoc.Save(stream);

                    // Begin async call to web request.

                    IAsyncResult asyncResult = webRequest.BeginGetResponse(null, null);

                    // Suspend this thread until call is complete. 
                    asyncResult.AsyncWaitHandle.WaitOne();

                    using (WebResponse webResponse = webRequest.EndGetResponse(asyncResult))
                    {
                        using (StreamReader rd = new StreamReader(webResponse.GetResponseStream()))
                        {
                            SOAPResult = rd.ReadToEnd();
                        }
                    }
                }
            }
            return SOAPResult;

        }
    }
}
