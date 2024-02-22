using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;

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
        EventLogger EventLog = new EventLogger("DomainReroutingAgent");

        static readonly string RegistryHive = @"Software\TransportAgents\DomainReroutingAgent";
        static readonly string RegistryKeyAgentEnabled = "AgentEnabled";
        static readonly string RegistryKeyDebugEnabled = "DebugEnabled";
        static readonly string RegistryKeyEnabledSendersToReroute = "SendersToReroute";

        static bool IsDebugEnabled = true;
        static bool IsAgentEnabled = true;
        static Dictionary<string, string> SenderOverrides = new Dictionary<string, string>();

        public DomainReroutingAgent_RoutingAgent()
        {
            base.OnSubmittedMessage += new SubmittedMessageEventHandler(AccessRegistryConfiguration);
            base.OnResolvedMessage += new ResolvedMessageEventHandler(SendViaCustomRoutingDomain);
            base.OnCategorizedMessage += new CategorizedMessageEventHandler(RemoveUnsupportedHeaders);
        }

        void AccessRegistryConfiguration(SubmittedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                RegistryKey registryPath = null;
                Stopwatch stopwatch = Stopwatch.StartNew();
                EventLog.AppendLogEntry(String.Format("Accessign Registry to check configuration of {0} in DomainReroutingAgent:AccessRegistryConfiguration", RegistryHive));

                registryPath = Registry.CurrentUser.OpenSubKey(RegistryHive, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);

                if (registryPath != null)
                {
                    EventLog.AppendLogEntry(String.Format("Registry Key {0} esists and contains {1} entries", RegistryHive, registryPath.ValueCount));

                    Boolean.TryParse(registryPath.GetValue(RegistryKeyAgentEnabled).ToString(), out IsAgentEnabled);
                    EventLog.AppendLogEntry(String.Format("The value of {0} is: {1}", RegistryKeyAgentEnabled, IsAgentEnabled));

                    Boolean.TryParse(registryPath.GetValue(RegistryKeyDebugEnabled).ToString(), out IsDebugEnabled);
                    EventLog.AppendLogEntry(String.Format("The value of {0} is: {1}", RegistryKeyDebugEnabled, IsDebugEnabled));

                    foreach (string s in (string[])registryPath.GetValue(RegistryKeyEnabledSendersToReroute))
                    {
                        string sender = s.Substring(0, s.IndexOf("|"));
                        string domain = s.Substring(s.IndexOf("|") + 1);
                        EventLog.AppendLogEntry(String.Format("Read override {0}:{1} from registry", sender, domain));

                        if (!SenderOverrides.ContainsKey(sender))
                        {
                            SenderOverrides.Add(sender, domain);
                            EventLog.AppendLogEntry(String.Format("Added to the list of overrides {0}:{1} to runtime configuration", sender, domain));
                        }
                    }
                }
                else
                {
                    EventLog.AppendLogEntry(String.Format("Registry Key {0} does not esist", RegistryHive));
                    Registry.CurrentUser.CreateSubKey(RegistryHive);
                    registryPath = Registry.CurrentUser.OpenSubKey(RegistryHive, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
                    registryPath.SetValue(RegistryKeyAgentEnabled, "True", RegistryValueKind.String);
                    registryPath.SetValue(RegistryKeyDebugEnabled, "True", RegistryValueKind.String);
                    registryPath.SetValue(RegistryKeyEnabledSendersToReroute, new[] { "noreply@toniolo.cloud|acs.toniolo.cloud" }, RegistryValueKind.MultiString);
                    EventLog.AppendLogEntry(String.Format("Created sample registry keys with test values"));
                }

                stopwatch.Stop();
                EventLog.AppendLogEntry(String.Format("AccessRegistryConfiguration took {0} ms to execute", stopwatch.ElapsedMilliseconds));
                EventLog.LogDebug(IsDebugEnabled);

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in DomainReroutingAgent:AccessRegistryConfiguration");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;
        }

        void SendViaCustomRoutingDomain(ResolvedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in DomainReroutingAgent:SendViaCustomRoutingDomain. IsAgentEnabled is set to: {3}", messageId, sender, subject, IsAgentEnabled));

                if (IsAgentEnabled && SenderOverrides.ContainsKey(sender))
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

                EventLog.AppendLogEntry(String.Format("SendViaCustomRoutingDomain took {0} ms to execute", stopwatch.ElapsedMilliseconds));
                EventLog.LogDebug(IsDebugEnabled);

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in DomainReroutingAgent:SendViaCustomRoutingDomain");
                EventLog.AppendLogEntry(ex);
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
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in DomainReroutingAgent:RemoveUnsupportedHeaders. IsAgentEnabled is set to: {3}", messageId, sender, subject, IsAgentEnabled));

                if (IsAgentEnabled && SenderOverrides.ContainsKey(sender))
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

                EventLog.AppendLogEntry(String.Format("RemoveUnsupportedHeaders took {0} ms to execute", stopwatch.ElapsedMilliseconds));
                EventLog.LogDebug(IsDebugEnabled);

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in DomainReroutingAgent:RemoveUnsupportedHeaders");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }

    }

}