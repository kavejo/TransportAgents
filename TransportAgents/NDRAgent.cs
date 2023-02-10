using System;
using System.Collections.Specialized;
using System.Configuration;
using System.Text;
using System.Diagnostics;
using System.IO;
using Microsoft.Exchange.Data.Globalization;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Exchange.Data.Common;
using Microsoft.Exchange.Data.Mime;



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

        public NDRAgent_Agent(SmtpServer server)
        {
            this.OnRoutedMessage += OnRoutedMessageBlockNDR;
        }

        private void OnRoutedMessageBlockNDR(RoutedMessageEventSource source, QueuedMessageEventArgs e)
        {
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
                using (TextLogger textLog = new TextLogger(@"F:\\Transport Agents", @"NDRAgent_Log.txt"))
                {
                    StringBuilder errorEntry = new StringBuilder();

                    errorEntry.AppendLine("------------------------------------------------------------");
                    errorEntry.AppendLine("EXCEPTION!!!");
                    errorEntry.AppendLine("------------------------------------------------------------");
                    errorEntry.AppendLine(String.Format("HResult: {0}", ex.HResult.ToString()));
                    errorEntry.AppendLine(String.Format("Message: {0}", ex.Message.ToString()));
                    errorEntry.AppendLine(String.Format("Source: {0}", ex.Source.ToString()));
                    errorEntry.AppendLine(String.Format("InnerException: {0}", ex.InnerException.ToString()));
                    errorEntry.AppendLine(String.Format("StackTrace: {0}", ex.StackTrace.ToString()));

                    textLog.WriteToText(errorEntry.ToString(), "Error");
                }
            }
        }
    }

}
