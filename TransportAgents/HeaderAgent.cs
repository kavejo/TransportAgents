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
    public class HeaderAgent : RoutingAgentFactory
    {

        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new HeaderAgent_Agent(server);
        }
    }

    public class HeaderAgent_Agent : RoutingAgent
    {

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
                using (TextLogger textLog = new TextLogger(@"F:\\Transport Agents", @"HeaderAgent_Log.txt"))
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

            return;

        }
    }

}
