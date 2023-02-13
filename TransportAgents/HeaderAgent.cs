using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Diagnostics;
using System.Text;



namespace TransportAgents
{
    public class HeaderAgent : RoutingAgentFactory
    {

        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new HeaderAgent_Agent(server);
        }
    }

    public class HeaderAgent_Agent : RoutingAgent
    {

        static string LogFile = String.Format("F:\\Transport Agents\\{0}.log", "HeaderAgent");
        TextLogger TextLog = new TextLogger(LogFile);

        public HeaderAgent_Agent(SmtpServer server)
        {
            this.OnRoutedMessage += OnRoutedMessageInsertHeader;
        }

        private void OnRoutedMessageInsertHeader(RoutedMessageEventSource source, QueuedMessageEventArgs e)
        {
            TextLog.WriteToText("Entering: OnRoutedMessageInsertHeader");

            string headerName = "X-TOTONI";
            string headerValue = "This message has been processed by a Transpor Agent written by TOTONI@MICROSOFT.COM";

            try
            {
                MimeDocument mimeDoc = e.MailItem.Message.MimeDocument;
                HeaderList headers = mimeDoc.RootPart.Headers;
                Header headerNamePresent = headers.FindFirst(headerName);

                if (headerNamePresent == null)
                {
                    MimeNode lastHeader = headers.LastChild;
                    TextHeader newHeader = new TextHeader(headerName, headerValue);
                    headers.InsertAfter(newHeader, lastHeader);
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
