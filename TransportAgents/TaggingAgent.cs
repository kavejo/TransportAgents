using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Smtp;
using System;
using System.Diagnostics;

namespace TransportAgents
{
    public class TaggingAgent : SmtpReceiveAgentFactory
    {
        public override SmtpReceiveAgent CreateAgent(SmtpServer server)
        {
            return new TaggingAgent_Agent();
        }
    }

    public class TaggingAgent_Agent : SmtpReceiveAgent
    {

        EventLogger EventLog = new EventLogger("TaggingAgent");
        static bool IsDebugEnabled = true;

        public TaggingAgent_Agent()
        {
            OnRcptCommand += StripTagOutOfAddress;
        }

        private void StripTagOutOfAddress(ReceiveCommandEventSource receiveMessageEventSource, RcptCommandEventArgs eventArgs)
        {

            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                EventLog.AppendLogEntry(String.Format("Entering: TaggingAgent:StripTagOutOfAddress"));

                RoutingAddress initialRecipientAddress = eventArgs.RecipientAddress;

                if (initialRecipientAddress.LocalPart.Contains("+"))
                {
                    int plusIndex = initialRecipientAddress.LocalPart.IndexOf("+");
                    string recipient = initialRecipientAddress.LocalPart.Substring(0, plusIndex);
                    string tag = initialRecipientAddress.LocalPart.Substring(plusIndex + 1);
                    string domain = initialRecipientAddress.DomainPart;
                    string revisedRecipientAddress = recipient + "@" + domain;
                    eventArgs.RecipientAddress = RoutingAddress.Parse(revisedRecipientAddress);
                    eventArgs.OriginalRecipient = revisedRecipientAddress;
                }

                EventLog.AppendLogEntry(String.Format("TaggingAgent:StripTagOutOfAddress took {0} ms to execute", stopwatch.ElapsedMilliseconds));
                EventLog.LogDebug(IsDebugEnabled);

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in TaggingAgent:StripTagOutOfAddress");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }
    }

}

