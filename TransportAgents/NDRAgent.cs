using Microsoft.Exchange.Data.Globalization;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;



namespace TransportAgents
{
    public class NDRAgent : RoutingAgentFactory
    {

        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new NDRAgent_Agent(server);
        }
    }

    public class NDRAgent_Agent : RoutingAgent
    {

        static string LogFile = String.Format("F:\\Transport Agents\\{0}.log", "NDRAgent");
        TextLogger TextLog = new TextLogger(LogFile);

        public NDRAgent_Agent(SmtpServer server)
        {
            this.OnRoutedMessage += OnRoutedMessageBlockNDR;
        }

        private void OnRoutedMessageBlockNDR(RoutedMessageEventSource source, QueuedMessageEventArgs e)
        {
            TextLog.WriteToText("Entering: OnRoutedMessageBlockNDR");

            try
            {
                Body body = e.MailItem.Message.Body;
                Encoding encoding = Charset.GetEncoding(e.MailItem.Message.Body.CharsetName);
                string bodyValue = String.Empty;

                using (StreamReader stream = new StreamReader(body.GetContentReadStream(), encoding, true))
                {
                    bodyValue = stream.ReadToEnd();
                    stream.Dispose();
                }

                if (e.MailItem.Message.IsSystemMessage == true && bodyValue.Contains("DELETE"))
                {
                    source.Delete("NDRRoutingAgent");
                }
            }
            catch (Exception ex)
            {
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText("EXCEPTION!!!");
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText(String.Format("HResult: {0}", ex.HResult.ToString()));
                TextLog.WriteToText(String.Format("Message: {0}", ex.Message.ToString()));
                TextLog.WriteToText(String.Format("Source: {0}", ex.Source.ToString()));
            }

            return;

        }
    }

}
