using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Xml.Linq;

namespace TransportAgents
{

    public class RecipientDomainRewriteAgent : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new RecipientDomainRewriteAgent_Agent();
        }
    }

    public class RecipientDomainRewriteAgent_Agent : RoutingAgent
    {

        EventLogger EventLog = new EventLogger("RecipientDomainRewriteAgent");
        static bool IsDebugEnabled = true;

        static readonly string oldDomainToDecommission = "contoso.com";
        static readonly string newDomainToUse = "tailspin.com";
        static readonly List<String> excludedFromRewrite = new List<String>(new String[] { "someone@testdomain.local", "someoneelse@testdomain.local"});


        public RecipientDomainRewriteAgent_Agent()
        {
            base.OnSubmittedMessage += new SubmittedMessageEventHandler(RecipientDomainRewrite_OnSubmittedMessage);
        }

        void RecipientDomainRewrite_OnSubmittedMessage(SubmittedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            
            Stopwatch stopwatch = Stopwatch.StartNew();
            EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in RecipientDomainRewriteAgent:RecipientDomainRewrite_OnSubmittedMessage", evtMessage.MailItem.Message.MessageId.ToString(), evtMessage.MailItem.FromAddress.ToString().Trim(), evtMessage.MailItem.Message.Subject.ToString().Trim()));

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
                EventLog.AppendLogEntry("Exception in RecipientDomainRewriteAgent:RecipientDomainRewrite_OnSubmittedMessage at Sender Exclusion Check");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            /////
            ///// P1 HEADER (SMTP envelope)
            /////
            try
            {
                for (int intCounter = evtMessage.MailItem.Recipients.Count - 1; intCounter >= 0; intCounter--)
                {

                    string msgRecipientP1 = evtMessage.MailItem.Recipients[intCounter].Address.LocalPart;
                    string msgDomainP1 = evtMessage.MailItem.Recipients[intCounter].Address.DomainPart;

                    EventLog.AppendLogEntry(String.Format("Evaluating Recipient P1: {0}@{1}", msgRecipientP1, msgDomainP1));

                    if (msgDomainP1.ToLower() == oldDomainToDecommission.ToLower())
                    {
                        evtMessage.MailItem.Recipients[intCounter].Address = new RoutingAddress(msgRecipientP1, newDomainToUse);
                        EventLog.AppendLogEntry(String.Format("Updated Recipient P1: {0}@{1} to {2}@{3} ", msgRecipientP1, msgDomainP1, msgRecipientP1, newDomainToUse));
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in RecipientDomainRewriteAgent:RecipientDomainRewrite_OnSubmittedMessage at P1 Rewrite");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            /////
            ///// P2 HEADER (Message envelope)
            /////
            try
            {
                for (int intCounter = evtMessage.MailItem.Message.To.Count - 1; intCounter >= 0; intCounter--)
                {
                    int atIndex = evtMessage.MailItem.Message.To[intCounter].SmtpAddress.IndexOf("@");
                    int recLength = evtMessage.MailItem.Message.To[intCounter].SmtpAddress.Length;
                    string msgRecipientP2 = evtMessage.MailItem.Message.To[intCounter].SmtpAddress.Substring(0, atIndex);
                    string msgDomainP2 = evtMessage.MailItem.Message.To[intCounter].SmtpAddress.Substring(atIndex + 1, recLength - atIndex - 1);

                    EventLog.AppendLogEntry(String.Format("Evaluating Recipient P2 in TO: {0}@{1}", msgRecipientP2, msgDomainP2));

                    if (msgDomainP2.ToLower() == oldDomainToDecommission.ToLower())
                    {
                        evtMessage.MailItem.Message.To[intCounter].SmtpAddress = msgRecipientP2 + "@" + newDomainToUse;
                        EventLog.AppendLogEntry(String.Format("Updated Recipient P2 in TO: {0}@{1} to {2}@{3} ", msgRecipientP2, msgDomainP2, msgRecipientP2, newDomainToUse));
                    }
                }

                for (int intCounter = evtMessage.MailItem.Message.Cc.Count - 1; intCounter >= 0; intCounter--)
                {
                    int atIndex = evtMessage.MailItem.Message.Cc[intCounter].SmtpAddress.IndexOf("@");
                    int recLength = evtMessage.MailItem.Message.Cc[intCounter].SmtpAddress.Length;
                    string msgRecipientP2 = evtMessage.MailItem.Message.Cc[intCounter].SmtpAddress.Substring(0, atIndex);
                    string msgDomainP2 = evtMessage.MailItem.Message.Cc[intCounter].SmtpAddress.Substring(atIndex + 1, recLength - atIndex - 1);

                    EventLog.AppendLogEntry(String.Format("Evaluating Recipient P2 in CC: {0}@{1}", msgRecipientP2, msgDomainP2));

                    if (msgDomainP2.ToLower() == oldDomainToDecommission.ToLower())
                    {
                        evtMessage.MailItem.Message.Cc[intCounter].SmtpAddress = msgRecipientP2 + "@" + newDomainToUse;
                        EventLog.AppendLogEntry(String.Format("Updated Recipient P2 in CC: {0}@{1} to {2}@{3} ", msgRecipientP2, msgDomainP2, msgRecipientP2, newDomainToUse));
                    }
                }

                for (int intCounter = evtMessage.MailItem.Message.Bcc.Count - 1; intCounter >= 0; intCounter--)
                {
                    int atIndex = evtMessage.MailItem.Message.Bcc[intCounter].SmtpAddress.IndexOf("@");
                    int recLength = evtMessage.MailItem.Message.Bcc[intCounter].SmtpAddress.Length;
                    string msgRecipientP2 = evtMessage.MailItem.Message.Bcc[intCounter].SmtpAddress.Substring(0, atIndex);
                    string msgDomainP2 = evtMessage.MailItem.Message.Bcc[intCounter].SmtpAddress.Substring(atIndex + 1, recLength - atIndex - 1);

                    EventLog.AppendLogEntry(String.Format("Evaluating Recipient P2 in BCC: {0}@{1}", msgRecipientP2, msgDomainP2));

                    if (msgDomainP2.ToLower() == oldDomainToDecommission.ToLower())
                    {
                        evtMessage.MailItem.Message.Bcc[intCounter].SmtpAddress = msgRecipientP2 + "@" + newDomainToUse;
                        EventLog.AppendLogEntry(String.Format("Updated Recipient P2 in BCC: {0}@{1} to {2}@{3} ", msgRecipientP2, msgDomainP2, msgRecipientP2, newDomainToUse));
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in RecipientDomainRewriteAgent:RecipientDomainRewrite_OnSubmittedMessage at P2 Rewrite");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            EventLog.AppendLogEntry(String.Format("RecipientDomainRewriteAgent:RecipientDomainRewrite_OnSubmittedMessage took {0} ms to execute", stopwatch.ElapsedMilliseconds));
            EventLog.LogDebug(IsDebugEnabled);

            return;

        }
    }
}
