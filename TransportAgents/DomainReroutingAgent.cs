using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Collections.Generic;

namespace TransportAgents
{
    public class DomainReroutingAgent : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new DomainReroutingAgent_RoutingAgent();
        }
    }

    public class DomainReroutingAgent_RoutingAgent : RoutingAgent
    {
        static bool IsDebugEnabled = true;
        EventLogger EventLog = new EventLogger("DomainReroutingAgent");

        static readonly Dictionary<string, string> SenderOverrides = new Dictionary<string, string>
        {
            { "noreply@toniolo.cloud", "acs.toniolo.cloud" },
            { "acs_test@toniolo.cloud", "acs.toniolo.cloud" }
        };

        public DomainReroutingAgent_RoutingAgent()
        {
            base.OnResolvedMessage += new ResolvedMessageEventHandler(SendViaCustomRoutingDomain);
            base.OnCategorizedMessage += new CategorizedMessageEventHandler(RemoveUnsupportedHeaders);
        }

        void SendViaCustomRoutingDomain(ResolvedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in SendViaCustomRoutingDomain", messageId, sender, subject));

                if (SenderOverrides.ContainsKey(sender))
                {
                    EventLog.AppendLogEntry(String.Format("Rerouting messages as the sender is {0}", sender));
                    foreach (EnvelopeRecipient recipient in evtMessage.MailItem.Recipients)
                    {
                        EventLog.AppendLogEntry(String.Format("Evaluating recipient {0}", recipient.Address.ToString()));
                        if (recipient.RecipientCategory == RecipientCategory.InDifferentOrganization || recipient.RecipientCategory == RecipientCategory.Unknown)
                        {
                            RoutingDomain customRoutingDomain = new RoutingDomain(SenderOverrides[sender]);
                            RoutingOverride destinationOverride = new RoutingOverride(customRoutingDomain, DeliveryQueueDomain.UseOverrideDomain);
                            source.SetRoutingOverride(recipient, destinationOverride);
                            EventLog.AppendLogEntry(String.Format("Recipient {0} overridden to {1} as it is {2}", recipient.Address.ToString(), SenderOverrides[sender], recipient.RecipientCategory));
                        }
                        else
                        {
                            EventLog.AppendLogEntry(String.Format("Recipient {0} not overridden as the recipient is {1}", recipient.Address.ToString(), recipient.RecipientCategory));
                        }
                    }
                }

                EventLog.LogDebug(IsDebugEnabled);

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in SendViaCustomRoutingDomain");
                EventLog.AppendLogEntry(ex.ToString());
                EventLog.LogError();
            }

            return;

        }

        void RemoveUnsupportedHeaders(CategorizedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                int externalRecipients = 0;

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in RemoveUnsupportedHeaders", messageId, sender, subject));

                if (SenderOverrides.ContainsKey(sender))
                {
                    EventLog.AppendLogEntry(String.Format("Evaluating headers as the sender is {0}", sender));

                    foreach (EnvelopeRecipient recipient in evtMessage.MailItem.Recipients)
                    {
                        if (recipient.RecipientCategory == RecipientCategory.InDifferentOrganization || recipient.RecipientCategory == RecipientCategory.Unknown)
                        {
                            externalRecipients++; 
                        }
                    }
                    EventLog.AppendLogEntry(String.Format("There are {0} external recipients", externalRecipients));

                    if (externalRecipients > 0)
                    {
                        EventLog.AppendLogEntry(String.Format("Removing empty headers as there are extenral recipients"));
                        foreach (Header header in evtMessage.MailItem.Message.MimeDocument.RootPart.Headers)
                        {
                            EventLog.AppendLogEntry(String.Format("Inspeceting header {0}: {1}", header.Name, String.IsNullOrEmpty(header.Value) ? String.Empty : header.Value));
                            if (header.Value == null || header.Value.Length == 0 || String.IsNullOrEmpty(header.Value))
                            {
                                if (header.Name.ToLower() != "from" && header.Name.ToLower() != "to" && header.Name.ToLower() != "cc" && header.Name.ToLower() != "bcc")
                                {
                                    evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.RemoveChild(header);
                                    EventLog.AppendLogEntry(String.Format("Header {0} REMOVED", header.Name));
                                }
                            }
                        }
                    }

                    evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Name", "DomainReroutingAgent"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                    evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Creator", "Tommaso Toniolo"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                    evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Contact", "https://aka.ms/totoni"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                    EventLog.AppendLogEntry(String.Format("Added Headers of type {0}", "X-TransportAgent-*"));

                }
                
                EventLog.LogDebug(IsDebugEnabled);

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in RemoveUnsupportedHeaders");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }

    }

}