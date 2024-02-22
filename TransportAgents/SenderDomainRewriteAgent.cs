using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace TransportAgents
{
    public class SenderDomainRewriteAgent : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new SenderDomainRewriteAgent_Agent();
        }
    }

    public class SenderDomainRewriteAgent_Agent : RoutingAgent
    {
        EventLogger EventLog = new EventLogger("SenderDomainRewriteAgent");
        static bool IsDebugEnabled = true;

        static readonly string oldDomainToDecommission = "contoso.com";
        static readonly string newDomainToUse = "tailspin.com";
        static readonly List<String> excludedFromRewrite = new List<String>(new String[] { "someone@testdomain.local", "someoneelse@testdomain.local" });

        public SenderDomainRewriteAgent_Agent()
        {
            base.OnSubmittedMessage += new SubmittedMessageEventHandler(SenderDomainRewrite_OnSubmittedMessage);
        }

        void SenderDomainRewrite_OnSubmittedMessage(SubmittedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {

            Stopwatch stopwatch = Stopwatch.StartNew();
            EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in SenderDomainRewriteAgent:SenderDomainRewrite_OnSubmittedMessage", evtMessage.MailItem.Message.MessageId.ToString(), evtMessage.MailItem.FromAddress.ToString().Trim(), evtMessage.MailItem.Message.Subject.ToString().Trim()));

            /////
            ///// As there is a need for not parsing the messages sent from specific senders (listed in List<String> excludedFromRewrite).
            ///// In this case the class will return to the caller before executing any of the rewrite-logic.
            /////
            try
            {
                if (excludedFromRewrite.Contains(evtMessage.MailItem.FromAddress.ToString().Trim()) || excludedFromRewrite.Contains(evtMessage.MailItem.Message.Sender.ToString().Trim()))
                {
                    EventLog.AppendLogEntry(String.Format("Avoid to process as the sender is: P1={0}/P2={1}", evtMessage.MailItem.FromAddress.ToString().Trim(), evtMessage.MailItem.Message.Sender.ToString().Trim()));
                    EventLog.LogDebug(IsDebugEnabled);
                    return;
                }
            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in SenderDomainRewriteAgent:SenderDomainRewrite_OnSubmittedMessage at Sender Exclusion Check");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            /////
            ///// P1 HEADER (SMTP envelope)
            /////
            try
            {
                string msgSenderP1 = evtMessage.MailItem.FromAddress.LocalPart;
                string msgDomainP1 = evtMessage.MailItem.FromAddress.DomainPart;

                EventLog.AppendLogEntry(String.Format("Evaluating Sender P1: {0}@{1}", msgSenderP1, msgDomainP1));

                if (msgDomainP1.ToLower() == oldDomainToDecommission.ToLower())
                {
                    evtMessage.MailItem.FromAddress = new RoutingAddress(msgSenderP1, newDomainToUse);
                    EventLog.AppendLogEntry(String.Format("Updated Sender P1: {0}@{1} to {2}@{3} ", msgSenderP1, msgDomainP1, msgSenderP1, newDomainToUse));
                }
            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in SenderDomainRewriteAgent:SenderDomainRewrite_OnSubmittedMessage at P1 Rewrite");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            /////
            ///// P2 HEADER (Message envelope)
            /////
            try
            {
                int atIndex = evtMessage.MailItem.Message.From.SmtpAddress.IndexOf("@");
                int recLength = evtMessage.MailItem.Message.From.SmtpAddress.Length;
                string msgSenderP2 = evtMessage.MailItem.Message.From.SmtpAddress.Substring(0, atIndex);
                string msgDomainP2 = evtMessage.MailItem.Message.From.SmtpAddress.Substring(atIndex + 1, recLength - atIndex - 1);

                EventLog.AppendLogEntry(String.Format("Evaluating Sender P2: {0}@{1}", msgSenderP2, msgDomainP2));

                if (msgDomainP2.ToLower() == oldDomainToDecommission.ToLower())
                {
                    evtMessage.MailItem.Message.From.SmtpAddress = msgSenderP2 + "@" + newDomainToUse;
                    evtMessage.MailItem.Message.Sender.SmtpAddress = msgSenderP2 + "@" + newDomainToUse;
                    EventLog.AppendLogEntry(String.Format("Updated Sender P2: {0}@{1} to {2}@{3} ", msgSenderP2, msgDomainP2, msgSenderP2, newDomainToUse));
                }
            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in SenderDomainRewriteAgent:SenderDomainRewrite_OnSubmittedMessage at P2 Rewrite");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            EventLog.AppendLogEntry(String.Format("SenderDomainRewriteAgent:SenderDomainRewrite_OnSubmittedMessage took {0} ms to execute", stopwatch.ElapsedMilliseconds));
            EventLog.LogDebug(IsDebugEnabled);

            return;

        }
    }
}
