using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

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
        static readonly string HeaderReroutingAgentTargetName = "X-HeaderReroutingAgent-Target";
        static string HeaderReroutingAgentTargetValue = String.Empty;
        static readonly string HeaderReroutingAgentP1P2MismatchActionName = "X-HeaderReroutingAgent-P1P2MismatchAction";
        static string HeaderReroutingAgentP1P2MismatchActionValue = String.Empty;
        static readonly string HeaderReroutingAgentForceP1Name = "X-HeaderReroutingAgent-ForceP1";
        static string HeaderReroutingAgentForceP1Value = String.Empty;
        static List<string> HeadersToRetain = new List<string>() { "From", "To", "Cc", "Bcc", "Subject", "Message-ID", "Content-Type", "Content-Transfer-Encoding", "MIME-Version", HeaderReroutingAgentTargetName, HeaderReroutingAgentP1P2MismatchActionName, HeaderReroutingAgentForceP1Name };
        static bool DebugEnabled = true;

        public HeaderReroutingAgent_RoutingAgent()
        {
            base.OnResolvedMessage += new ResolvedMessageEventHandler(OverrideRoutingDomain);
            base.OnRoutedMessage += new RoutedMessageEventHandler(OverrideSenderAddress);
            base.OnCategorizedMessage += new CategorizedMessageEventHandler(RemoveAllHeaders);
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

                Header HeaderReroutingAgentTarget = headers.FindFirst(HeaderReroutingAgentTargetName);
                
                if (HeaderReroutingAgentTarget != null)
                {
                    EventLog.AppendLogEntry(String.Format("Rerouting messages as the control header {0} is present", HeaderReroutingAgentTargetName));
                    HeaderReroutingAgentTargetValue = HeaderReroutingAgentTarget.Value.Trim();

                    if (!String.IsNullOrEmpty(HeaderReroutingAgentTargetValue) && (Uri.CheckHostName(HeaderReroutingAgentTargetValue) == UriHostNameType.Dns))
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

        void OverrideSenderAddress(RoutedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                bool warningOccurred = false;
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                string P1Sender = evtMessage.MailItem.FromAddress.ToString();
                string P2Sender = evtMessage.MailItem.Message.Sender.SmtpAddress;
                HeaderList headers = evtMessage.MailItem.Message.MimeDocument.RootPart.Headers;
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in HeaderReroutingAgent:OverrideSenderAddress", messageId, sender, subject));

                Header HeaderReroutingAgentP1P2MismatchAction = headers.FindFirst(HeaderReroutingAgentP1P2MismatchActionName);

                if (HeaderReroutingAgentP1P2MismatchAction != null)
                {
                    EventLog.AppendLogEntry(String.Format("Evaluating P1/P2 Sender Mismatch as the control header {0} is present", HeaderReroutingAgentP1P2MismatchActionName));
                    HeaderReroutingAgentP1P2MismatchActionValue = HeaderReroutingAgentP1P2MismatchAction.Value.Trim().ToUpper();

                    if (!String.IsNullOrEmpty(HeaderReroutingAgentP1P2MismatchActionValue) &&
                        (String.Equals(HeaderReroutingAgentP1P2MismatchActionValue, "UseP1", StringComparison.OrdinalIgnoreCase) || 
                         String.Equals(HeaderReroutingAgentP1P2MismatchActionValue, "UseP2", StringComparison.OrdinalIgnoreCase) || 
                         String.Equals(HeaderReroutingAgentP1P2MismatchActionValue, "None",  StringComparison.OrdinalIgnoreCase)  )
                    )
                    {
                        EventLog.AppendLogEntry(String.Format("P1/P2 Mismatch Action is valid as the header {0} is set to {1}", HeaderReroutingAgentP1P2MismatchActionName, HeaderReroutingAgentP1P2MismatchActionValue));

                        EventLog.AppendLogEntry(String.Format("P1 Sender is set to: {0}", P1Sender));
                        EventLog.AppendLogEntry(String.Format("P2 Sender is set to: {0}", P2Sender));

                        switch (HeaderReroutingAgentP1P2MismatchActionValue)
                        {
                            case "USEP1":
                                evtMessage.MailItem.Message.Sender.SmtpAddress = P1Sender;
                                evtMessage.MailItem.Message.From.SmtpAddress = P1Sender;
                                EventLog.AppendLogEntry(String.Format("P2 Sender has been set to: {0}", P1Sender));
                                break;
                            case "USEP2":
                                evtMessage.MailItem.FromAddress = new RoutingAddress(P2Sender);
                                EventLog.AppendLogEntry(String.Format("P1 Sender has been set to: {0}", P2Sender));
                                break;
                            case "NONE":
                                EventLog.AppendLogEntry(String.Format("No action has been taken as the header is set to {0}", HeaderReroutingAgentP1P2MismatchActionValue));
                                break;
                            default:
                                EventLog.AppendLogEntry(String.Format("P1 and P2 have been left unmodified"));
                                break;
                        }
                    }
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("There was a problem processing the {0} header value", HeaderReroutingAgentP1P2MismatchActionName));
                        EventLog.AppendLogEntry(String.Format("There value retrieved is: {0}; Valid (case insensitive) values are UseP1, UseP2, None", HeaderReroutingAgentP1P2MismatchActionValue));
                        warningOccurred = true;
                    }

                }

                Header HeaderReroutingAgentForceP1 = headers.FindFirst(HeaderReroutingAgentForceP1Name);

                if (HeaderReroutingAgentForceP1 != null)
                {
                    EventLog.AppendLogEntry(String.Format("Overriding P1 Sender as the control header {0} is present", HeaderReroutingAgentForceP1Name));
                    HeaderReroutingAgentForceP1Value = HeaderReroutingAgentForceP1.Value.Trim().ToUpper();

                    RoutingAddress newP1 = new RoutingAddress(HeaderReroutingAgentForceP1Value);
                    EventLog.AppendLogEntry(String.Format("The new P1 Sender to be forced is {0}", newP1.ToString()));

                    EventLog.AppendLogEntry(String.Format("P1 Sender is currently set to: {0}", P1Sender));
                    EventLog.AppendLogEntry(String.Format("P2 Sender is currently set to: {0}", P2Sender));

                    evtMessage.MailItem.FromAddress = newP1;
                    EventLog.AppendLogEntry(String.Format("Forced P1 Sender to {0}", evtMessage.MailItem.FromAddress.ToString()));
                }

                EventLog.AppendLogEntry(String.Format("HeaderReroutingAgent:OverrideSenderAddress took {0} ms to execute", stopwatch.ElapsedMilliseconds));

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
                EventLog.AppendLogEntry("Exception in HeaderReroutingAgent:OverrideSenderAddress");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }

        void RemoveAllHeaders(CategorizedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                HeaderList headers = evtMessage.MailItem.Message.MimeDocument.RootPart.Headers;
                int externalRecipients = 0;
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in HeaderReroutingAgent:RemoveAllHeaders", messageId, sender, subject));

                Header HeaderReroutingAgentTarget = headers.FindFirst(HeaderReroutingAgentTargetName);

                if (HeaderReroutingAgentTarget != null)
                {
                    EventLog.AppendLogEntry(String.Format("Removing All Headers as the control header {0} is present", HeaderReroutingAgentTargetName));

                    foreach (EnvelopeRecipient recipient in evtMessage.MailItem.Recipients)
                    {
                        EventLog.AppendLogEntry(String.Format("Evaluating recipient {0}@{1} which is {2}", recipient.Address.LocalPart.ToLower(), recipient.Address.DomainPart.ToLower(), recipient.RecipientCategory));
                        if (recipient.RecipientCategory != RecipientCategory.InSameOrganization)
                        {
                            externalRecipients++;
                        }
                    }
                    EventLog.AppendLogEntry(String.Format("There are {0} external recipient(s)", externalRecipients));

                    if (externalRecipients > 0)
                    {
                        EventLog.AppendLogEntry(String.Format("Removing all headers as there are external recipients"));
                        foreach (Header header in evtMessage.MailItem.Message.MimeDocument.RootPart.Headers)
                        {
                            if (HeadersToRetain.Contains( header.Name.ToLower(), StringComparer.OrdinalIgnoreCase) )
                            {
                                EventLog.AppendLogEntry(String.Format("KEPT header {0}: {1}", header.Name, String.IsNullOrEmpty(header.Value) ? String.Empty : header.Value));
                            }
                            else
                            {
                                evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.RemoveChild(header);
                                EventLog.AppendLogEntry(String.Format("REMOVED header {0}: {1}", header.Name, String.IsNullOrEmpty(header.Value) ? String.Empty : header.Value));
                            }
                        }
                    }

                    evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Name", "HeaderReroutingAgent"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                    evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Creator", "Tommaso Toniolo"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                    evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Contact", "https://aka.ms/totoni"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                    EventLog.AppendLogEntry(String.Format("Added Headers of type {0}", "X-TransportAgent-*"));

                }

                EventLog.AppendLogEntry(String.Format("HeaderReroutingAgent:RemoveAllHeaders took {0} ms to execute", stopwatch.ElapsedMilliseconds));
                EventLog.LogDebug(DebugEnabled);

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in HeaderReroutingAgent:RemoveAllHeaders");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }

    }
}
