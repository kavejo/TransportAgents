using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Diagnostics;

namespace TransportAgents
{
    public  class HeaderReroutingAgent : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new HeaderReroutingAgent_RoutingAgent();
        }
    }

    public class HeaderReroutingAgent_RoutingAgent : RoutingAgent
    {
        EventLogger EventLog = new EventLogger("HeaderReroutingAgent");
        static readonly string HeaderReroutingAgentEnabledName = "X-HeaderReroutingAgent-Enabled";
        static bool HeaderReroutingAgentEnabledValue = false;
        static readonly string HeaderReroutingAgentTargetName = "X-HeaderReroutingAgent-Target";
        static string HeaderReroutingAgentTargetValue = String.Empty;
        static bool DebugEnabled = true;

        public HeaderReroutingAgent_RoutingAgent()
        {
            base.OnResolvedMessage += new ResolvedMessageEventHandler(OverrideRoutingDomain);
            base.OnCategorizedMessage += new CategorizedMessageEventHandler(RemoveEmptyHeaders);
        }

        void OverrideRoutingDomain(ResolvedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                bool warningOccurred = false;
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                HeaderList headers = evtMessage.MailItem.Message.MimeDocument.RootPart.Headers;
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in HeaderReroutingAgent:OverrideRoutingDomain", messageId, sender, subject));

                Header HeaderReroutingAgentEnabled = headers.FindFirst(HeaderReroutingAgentEnabledName);
                bool valueConversionResult = Boolean.TryParse(HeaderReroutingAgentEnabled.Value, out HeaderReroutingAgentEnabledValue);


                if (HeaderReroutingAgentEnabledValue)
                {
                    EventLog.AppendLogEntry(String.Format("Rerouting messages as the control header {0} is set to {1}", HeaderReroutingAgentEnabledName, HeaderReroutingAgentEnabledValue));
                    Header HeaderReroutingAgentTarget = headers.FindFirst(HeaderReroutingAgentTargetName);
                    HeaderReroutingAgentTargetValue = HeaderReroutingAgentTarget.Value.Trim();

                    if (String.IsNullOrEmpty(HeaderReroutingAgentTargetValue) && (Uri.CheckHostName(HeaderReroutingAgentTargetValue) == UriHostNameType.Dns))
                    {
                        EventLog.AppendLogEntry(String.Format("Rerouting domain is valid as the header {0} is set to {1}", HeaderReroutingAgentTargetName, HeaderReroutingAgentTargetValue));

                        foreach (EnvelopeRecipient recipient in evtMessage.MailItem.Recipients)
                        {
                            EventLog.AppendLogEntry(String.Format("Evaluating recipient {0}", recipient.Address.ToString()));
                            if (recipient.RecipientCategory == RecipientCategory.InSameOrganization)
                            {
                                EventLog.AppendLogEntry(String.Format("Recipient {0} not overridden as the recipient domain IS internal; the recipient is categorzed as {1}", recipient.Address.ToString(), recipient.RecipientCategory));
                            }
                            else
                            {
                                RoutingDomain customRoutingDomain = new RoutingDomain(HeaderReroutingAgentTargetValue);
                                RoutingOverride destinationOverride = new RoutingOverride(customRoutingDomain, DeliveryQueueDomain.UseOverrideDomain);
                                source.SetRoutingOverride(recipient, destinationOverride);
                                EventLog.AppendLogEntry(String.Format("Recipient {0} overridden to {1} as the recipient domain IS NOT internal; the recipient is categorzed as {2}", recipient.Address.ToString(), HeaderReroutingAgentTargetValue, recipient.RecipientCategory));
                            }
                        }
                    }
                    else 
                    {
                        EventLog.AppendLogEntry(String.Format("There was a problem processing the {0} header value", HeaderReroutingAgentTargetName));
                        EventLog.AppendLogEntry(String.Format("There value retrieved is: {0}", HeaderReroutingAgentTargetValue));
                        warningOccurred = true;
                    }

                }

                EventLog.AppendLogEntry(String.Format("HeaderReroutingAgent:OverrideRoutingDomain took {0} ms to execute", stopwatch.ElapsedMilliseconds));

                if (warningOccurred)
                {
                    EventLog.LogWarning();
                }
                else
                {
                    EventLog.LogDebug(DebugEnabled);
                }

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in HeaderReroutingAgent:OverrideRoutingDomain");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }

        void RemoveEmptyHeaders(CategorizedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                HeaderList headers = evtMessage.MailItem.Message.MimeDocument.RootPart.Headers;
                int externalRecipients = 0;
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in HeaderReroutingAgent:RemoveEmptyHeaders", messageId, sender, subject));

                Header HeaderReroutingAgentEnabled = headers.FindFirst(HeaderReroutingAgentEnabledName);
                bool valueConversionResult = Boolean.TryParse(HeaderReroutingAgentEnabled.Value, out HeaderReroutingAgentEnabledValue);

                if (HeaderReroutingAgentEnabledValue)
                {
                    EventLog.AppendLogEntry(String.Format("Evaluating message headers as the control header {0} is set to {1}", HeaderReroutingAgentEnabledName, HeaderReroutingAgentEnabledValue));

                    foreach (EnvelopeRecipient recipient in evtMessage.MailItem.Recipients)
                    {
                        EventLog.AppendLogEntry(String.Format("Evaluating recipient {0}@{1} which is {2}", recipient.Address.LocalPart.ToLower(), recipient.Address.DomainPart.ToLower(), recipient.RecipientCategory));
                        if (recipient.RecipientCategory != RecipientCategory.InSameOrganization)
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

                EventLog.AppendLogEntry(String.Format("HeaderReroutingAgent:RemoveEmptyHeaders took {0} ms to execute", stopwatch.ElapsedMilliseconds));
                EventLog.LogDebug(DebugEnabled);

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in HeaderReroutingAgent:RemoveEmptyHeaders");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }

    }
}
