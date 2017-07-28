using System;
using System.Collections.Generic;
using Apache.NMS;
using Apache.NMS.Util;
using System.Diagnostics;
using System.Threading;

namespace AutoTestLib.MessageEngine
{
    public class ActiveMQ
    {
        protected static AutoResetEvent semaphore = new AutoResetEvent(false);
        protected static ITextMessage message = null;
        protected static TimeSpan receiveTimeout = TimeSpan.FromSeconds(10);
        public static List<ITextMessage> listMessage = new List<ITextMessage>();

        public static void debugSample()
        {
            string ConnectUri = "activemq:tcp://csmsdc-vpampt01:61616";
            string QueueName = "AMPLAINV.PERFADJNRT.OUT.COPY.1";
            string Username = "admin";
            string Password = "admin";

            List<ITextMessage> message = new List<ITextMessage>();

            message = ReadMessage(ConnectUri, QueueName, Username, Password);


            //Console.WriteLine("TextMessage.NMSMessageId: " + message[0].NMSMessageId);
            //for (int i = 0; i < message.Count(); i ++ )
            foreach (ITextMessage TextMessage in message)
            {
                Debug.WriteLine("TextMessage.NMSMessageId: " + TextMessage.NMSMessageId);
                Debug.WriteLine("TextMessage.Text: " + TextMessage.Text.ToString());
            }

        }


        public static List<ITextMessage> ReadMessage(string ConnectUri, string QueueName, string UserName, string Password)
        {
            //List<ITextMessage> listMessage = new List<ITextMessage>();

            Uri ConnectURI = new Uri(ConnectUri);
            IConnectionFactory factory = new NMSConnectionFactory(ConnectURI);

            using (IConnection connection = factory.CreateConnection(UserName, Password))
            using (ISession session = connection.CreateSession())
            {
                IDestination destination = SessionUtil.GetDestination(session, QueueName);
                using (IMessageConsumer consumer = session.CreateConsumer(destination))
                {
                    connection.Start();

                    consumer.Listener += new MessageListener(OnMessage);

                    // Wait for the message
                    semaphore.WaitOne((int)receiveTimeout.TotalMilliseconds, true);

                    /*
                    ITextMessage message = consumer.ReceiveNoWait() as ITextMessage;
                    while (message != null)
                    {
                        listMessage.Add(message);
                        message = consumer.Receive() as ITextMessage;
                    }
                     */
                }
            }

            return listMessage;
        }

        protected static void OnMessage(IMessage receivedMsg)
        {
            message = receivedMsg as ITextMessage;

            listMessage.Add(message);

            semaphore.Set();
        }
    }
}
