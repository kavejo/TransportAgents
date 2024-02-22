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

        EventLogger EventLog = new EventLogger("NDRAgent");
        static bool IsDebugEnabled = true;

        public NDRAgent_Agent(SmtpServer server)
        {
            this.OnRoutedMessage += OnRoutedMessageBlockNDR;
        }

        private void OnRoutedMessageBlockNDR(RoutedMessageEventSource source, QueuedMessageEventArgs e)
        {

            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                EventLog.AppendLogEntry(String.Format("Entering: NDRAgent:OnRoutedMessageBlockNDR"));

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

                EventLog.AppendLogEntry(String.Format("NDRAgent:OnRoutedMessageBlockNDR took {0} ms to execute", stopwatch.ElapsedMilliseconds));
                EventLog.LogDebug(IsDebugEnabled);
            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in NDRAgent:OnRoutedMessageBlockNDR");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }
    }

}
