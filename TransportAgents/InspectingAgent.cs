using Microsoft.Exchange.Data.Mime;
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
    internal class InspectingAgent : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new InspectingAgent_RoutingAgent();
        }
    }

    public class InspectingAgent_RoutingAgent : RoutingAgent
    {
        EventLogger EventLog = new EventLogger("InspectingAgent");
        static bool DebugEnabled = true;

        public InspectingAgent_RoutingAgent()
        {
            base.OnResolvedMessage += new ResolvedMessageEventHandler(WriteMessageProperitesOnLog);
        }

        void WriteMessageProperitesOnLog(ResolvedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processingmessage in InspectingAgent:WriteMessageProperitesOnLog"));

                EventLog.AppendLogEntry("==================== ENVELOPE - P1 ====================");
                EventLog.AppendLogEntry(String.Format("EnvelopeId: {0}", evtMessage.MailItem.EnvelopeId));
                EventLog.AppendLogEntry(String.Format("P1 Sender: {0}", evtMessage.MailItem.FromAddress.ToString().ToLower().Trim()));
                foreach (var recipient in evtMessage.MailItem.Recipients)
                    EventLog.AppendLogEntry(String.Format("P1 Recipient: {0}", recipient.Address.ToString().ToLower()));
                EventLog.AppendLogEntry(String.Format("IsSystemMessage: {0}", evtMessage.MailItem.Message.IsSystemMessage));
                EventLog.AppendLogEntry(String.Format("IsInterpersonalMessage: {0}", evtMessage.MailItem.Message.IsInterpersonalMessage));
                EventLog.AppendLogEntry(String.Format("IsOpaqueMessage: {0}", evtMessage.MailItem.Message.IsOpaqueMessage));
                EventLog.AppendLogEntry(String.Format("OriginatingDomain: {0}", evtMessage.MailItem.OriginatingDomain));
                EventLog.AppendLogEntry(String.Format("OriginatorOrganization: {0}", evtMessage.MailItem.OriginatorOrganization));
                EventLog.AppendLogEntry(String.Format("OriginalAuthenticator: {0}", evtMessage.MailItem.OriginalAuthenticator));
                foreach (var item in evtMessage.MailItem.Properties)
                    EventLog.AppendLogEntry(String.Format("Property - {0}: {1}", item.Key.ToString(), item.Value.ToString()));

                EventLog.AppendLogEntry("==================== HEADERS ====================");
                foreach (var header in evtMessage.MailItem.Message.MimeDocument.RootPart.Headers)
                    EventLog.AppendLogEntry(String.Format("{0}: {1}", header.Name, String.IsNullOrEmpty(header.Value) ? String.Empty : header.Value));

                EventLog.AppendLogEntry("==================== MESSAGE - P2 ====================");
                EventLog.AppendLogEntry(String.Format("MessageId: {0}", evtMessage.MailItem.Message.MessageId.ToString()));
                EventLog.AppendLogEntry(String.Format("Subject: {0}", evtMessage.MailItem.Message.Subject.Trim()));
                EventLog.AppendLogEntry(String.Format("P2 Sender: {0}", evtMessage.MailItem.Message.Sender.SmtpAddress.ToString().ToLower().Trim()));
                EventLog.AppendLogEntry(String.Format("P2 From: {0}", evtMessage.MailItem.Message.From.SmtpAddress.ToString().ToLower().Trim()));
                EventLog.AppendLogEntry(String.Format("MapiMessageClass: {0}", evtMessage.MailItem.Message.MapiMessageClass.ToString().Trim()));
                foreach (var recipient in evtMessage.MailItem.Message.To)
                    EventLog.AppendLogEntry(String.Format("P2 To: {0}", recipient.SmtpAddress.ToString().ToLower().Trim()));
                foreach (var recipient in evtMessage.MailItem.Message.Cc)
                    EventLog.AppendLogEntry(String.Format("P2 Cc: {0}", recipient.SmtpAddress.ToString().ToLower().Trim()));
                foreach (var recipient in evtMessage.MailItem.Message.Bcc)
                    EventLog.AppendLogEntry(String.Format("P2 Bcc: {0}", recipient.SmtpAddress.ToString().ToLower().Trim()));
                foreach (var recipient in evtMessage.MailItem.Message.ReplyTo)
                    EventLog.AppendLogEntry(String.Format("P2 ReplyTo: {0}", recipient.SmtpAddress.ToString().ToLower().Trim()));

                EventLog.AppendLogEntry(String.Format("InspectingAgent:WriteMessageProperitesOnLog took {0} ms to execute", stopwatch.ElapsedMilliseconds));

                if ( (evtMessage.MailItem.FromAddress.ToString().ToLower().Trim() != evtMessage.MailItem.Message.Sender.SmtpAddress.ToString().ToLower().Trim()) ||
                     (evtMessage.MailItem.FromAddress.ToString().ToLower().Trim() != evtMessage.MailItem.Message.From.SmtpAddress.ToString().ToLower().Trim()) )
                {
                    EventLog.AppendLogEntry("==================== IMPORTANT ====================");
                    EventLog.AppendLogEntry("Note that the P1 Sender and the P2 Sender mismatch. This can be source of problems");
                    EventLog.LogWarning();
                }
                else
                {
                    EventLog.LogDebug(DebugEnabled);
                }

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in InspectingAgent:WriteMessageProperitesOnLog");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }
    }
}
