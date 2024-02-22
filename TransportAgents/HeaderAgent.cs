using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Diagnostics;



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

        EventLogger EventLog = new EventLogger("HeaderAgent");
        static bool IsDebugEnabled = true;

        public HeaderAgent_Agent(SmtpServer server)
        {
            this.OnRoutedMessage += OnRoutedMessageInsertHeader;
        }

        private void OnRoutedMessageInsertHeader(RoutedMessageEventSource source, QueuedMessageEventArgs e)
        {

            string headerName = "X-TOTONI";
            string headerValue = "This message has been processed by a Transpor Agent written by TOTONI@MICROSOFT.COM";

            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                EventLog.AppendLogEntry(String.Format("Entering: HeaderAgent:OnRoutedMessageInsertHeader"));

                MimeDocument mimeDoc = e.MailItem.Message.MimeDocument;
                HeaderList headers = mimeDoc.RootPart.Headers;
                Header headerNamePresent = headers.FindFirst(headerName);

                if (headerNamePresent == null)
                {
                    MimeNode lastHeader = headers.LastChild;
                    TextHeader newHeader = new TextHeader(headerName, headerValue);
                    headers.InsertAfter(newHeader, lastHeader);
                }
                
                EventLog.AppendLogEntry(String.Format("HeaderAgent:OnRoutedMessageInsertHeader took {0} ms to execute", stopwatch.ElapsedMilliseconds));
                EventLog.LogDebug(IsDebugEnabled);

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in HeaderAgent:OnRoutedMessageInsertHeader");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }
    }

}
