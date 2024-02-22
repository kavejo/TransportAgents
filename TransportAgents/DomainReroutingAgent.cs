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
        static readonly string RegistryKeyOverrideRoutingDomainAgentEnabled = "OverrideRoutingDomain-AgentEnabled";
        static readonly string RegistryKeyOverrideRoutingDomainDebugEnabled = "OverrideRoutingDomain-DebugEnabled";
        static readonly string RegistryKeyRemoveEmptyHeadersAgentEnabled = "RemoveEmptyHeaders-AgentEnabled";
        static readonly string RegistryKeyRemoveEmptyHeadersDebugEnabled = "RemoveEmptyHeaders-DebugEnabled";

        static readonly string RegistryKeyEnabledSendersToReroute = "SendersToReroute";

        static bool OverrideRoutingDomainAgentEnabled = true; 
        static bool OverrideRoutingDomainDebugEnabled = true;
        static bool RemoveEmptyHeadersAgentEnabled = true;
        static bool RemoveEmptyHeadersDebugEnabled = true;
        static Dictionary<string, string> SenderOverrides = new Dictionary<string, string>();

        public DomainReroutingAgent_RoutingAgent()
        {
            base.OnSubmittedMessage += new SubmittedMessageEventHandler(RetrieveConfiguration);
            base.OnResolvedMessage += new ResolvedMessageEventHandler(OverrideRoutingDomain);
            base.OnCategorizedMessage += new CategorizedMessageEventHandler(RemoveEmptyHeaders);
        }

        void RetrieveConfiguration(SubmittedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                RegistryKey registryPath = null;
                Stopwatch stopwatch = Stopwatch.StartNew();
                EventLog.AppendLogEntry(String.Format("Accessign Registry to check configuration of {0} in DomainReroutingAgent:RetrieveConfiguration", RegistryHive));

                registryPath = Registry.CurrentUser.OpenSubKey(RegistryHive, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
                bool configurationReviewNecessary = false;
                bool valueConversionResult = true;

                if (registryPath != null)
                {
                    EventLog.AppendLogEntry(String.Format("Registry Key {0} esists and contains {1} entries", RegistryHive, registryPath.ValueCount));
                    if (registryPath.ValueCount != 5)
                    {
                        configurationReviewNecessary = true;
                    }

                    valueConversionResult = Boolean.TryParse(registryPath.GetValue(RegistryKeyOverrideRoutingDomainAgentEnabled).ToString(), out OverrideRoutingDomainAgentEnabled);
                    EventLog.AppendLogEntry(String.Format("The value of {0} is: {1}. Converting the registry value successfully: {2}", RegistryKeyOverrideRoutingDomainAgentEnabled, OverrideRoutingDomainAgentEnabled, valueConversionResult));
                    if (valueConversionResult == false)
                    {
                        configurationReviewNecessary = true;
                        EventLog.AppendLogEntry(String.Format("The registry key {0} is missing", RegistryKeyOverrideRoutingDomainAgentEnabled));
                    }

                    valueConversionResult = Boolean.TryParse(registryPath.GetValue(RegistryKeyOverrideRoutingDomainDebugEnabled).ToString(), out OverrideRoutingDomainDebugEnabled);
                    EventLog.AppendLogEntry(String.Format("The value of {0} is: {1}. Converting the registry value successfully: {2}", RegistryKeyOverrideRoutingDomainDebugEnabled, OverrideRoutingDomainDebugEnabled, valueConversionResult));
                    if (valueConversionResult == false)
                    {
                        configurationReviewNecessary = true;
                        EventLog.AppendLogEntry(String.Format("The registry key {0} is missing", RegistryKeyOverrideRoutingDomainDebugEnabled));
                    }

                    valueConversionResult = Boolean.TryParse(registryPath.GetValue(RegistryKeyRemoveEmptyHeadersAgentEnabled).ToString(), out RemoveEmptyHeadersAgentEnabled);
                    EventLog.AppendLogEntry(String.Format("The value of {0} is: {1}. Converting the registry value successfully: {2}", RegistryKeyRemoveEmptyHeadersAgentEnabled, RemoveEmptyHeadersAgentEnabled, valueConversionResult));
                    if (valueConversionResult == false)
                    {
                        configurationReviewNecessary = true;
                        EventLog.AppendLogEntry(String.Format("The registry key {0} is missing", RegistryKeyRemoveEmptyHeadersAgentEnabled));
                    }

                    valueConversionResult = Boolean.TryParse(registryPath.GetValue(RegistryKeyRemoveEmptyHeadersDebugEnabled).ToString(), out RemoveEmptyHeadersDebugEnabled);
                    EventLog.AppendLogEntry(String.Format("The value of {0} is: {1}. Converting the registry value successfully: {2}", RegistryKeyRemoveEmptyHeadersDebugEnabled, RemoveEmptyHeadersDebugEnabled, valueConversionResult));
                    if (valueConversionResult == false)
                    {
                        configurationReviewNecessary = true;
                        EventLog.AppendLogEntry(String.Format("The registry key {0} is missing", RegistryKeyRemoveEmptyHeadersDebugEnabled));
                    }

                    string[] retrievedOverrides = (string[])registryPath.GetValue(RegistryKeyEnabledSendersToReroute);
                    if (retrievedOverrides != null && retrievedOverrides.Length > 0)
                    {
                        foreach (string s in retrievedOverrides)
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
                        configurationReviewNecessary = true;
                        EventLog.AppendLogEntry(String.Format("The registry key {0} is missing", RegistryKeyEnabledSendersToReroute));
                    }
                }
                else
                {
                    configurationReviewNecessary = true;
                    EventLog.AppendLogEntry(String.Format("Registry Key {0} does not esist", RegistryHive));
                    Registry.CurrentUser.CreateSubKey(RegistryHive);
                    registryPath = Registry.CurrentUser.OpenSubKey(RegistryHive, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
                    registryPath.SetValue(RegistryKeyOverrideRoutingDomainAgentEnabled, "True", RegistryValueKind.String);
                    registryPath.SetValue(RegistryKeyOverrideRoutingDomainDebugEnabled, "True", RegistryValueKind.String);
                    registryPath.SetValue(RegistryKeyRemoveEmptyHeadersAgentEnabled, "True", RegistryValueKind.String);
                    registryPath.SetValue(RegistryKeyRemoveEmptyHeadersDebugEnabled, "True", RegistryValueKind.String);
                    registryPath.SetValue(RegistryKeyEnabledSendersToReroute, new[] { "noreply@toniolo.cloud|acs.toniolo.cloud" }, RegistryValueKind.MultiString);
                    EventLog.AppendLogEntry(String.Format("Created sample registry keys with test values"));
                }

                stopwatch.Stop();
                EventLog.AppendLogEntry(String.Format("DomainReroutingAgent:RetrieveConfiguration took {0} ms to execute", stopwatch.ElapsedMilliseconds));
                if(configurationReviewNecessary)
                {
                    EventLog.AppendLogEntry(String.Format(@"The configuration in HKEY_CURRENT_USER:\{0} appears incomplete. The following entries should be present.", RegistryHive));
                    EventLog.AppendLogEntry(String.Format("{0}: {1} of type REG_SZ", RegistryKeyOverrideRoutingDomainAgentEnabled, "True or False"));
                    EventLog.AppendLogEntry(String.Format("{0}: {1} of type REG_SZ", RegistryKeyOverrideRoutingDomainDebugEnabled, "True or False"));
                    EventLog.AppendLogEntry(String.Format("{0}: {1} of type REG_SZ", RegistryKeyRemoveEmptyHeadersAgentEnabled, "True or False"));
                    EventLog.AppendLogEntry(String.Format("{0}: {1} of type REG_SZ", RegistryKeyRemoveEmptyHeadersDebugEnabled, "True or False"));
                    EventLog.AppendLogEntry(String.Format("{0}: {1} of type REG_MULTI_SZ, one mapping per line", RegistryKeyEnabledSendersToReroute, "<user@domain.com>|<domain.tld>"));
                    EventLog.LogWarning();
                }
                else
                {
                    EventLog.LogDebug(OverrideRoutingDomainDebugEnabled);
                }
                

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in DomainReroutingAgent:RetrieveConfiguration");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;
        }

        void OverrideRoutingDomain(ResolvedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in DomainReroutingAgent:OverrideRoutingDomain. OverrideRoutingDomainAgentEnabled is set to: {3}", messageId, sender, subject, OverrideRoutingDomainAgentEnabled));

                if (OverrideRoutingDomainAgentEnabled && SenderOverrides.ContainsKey(sender))
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

                EventLog.AppendLogEntry(String.Format("DomainReroutingAgent:OverrideRoutingDomain took {0} ms to execute", stopwatch.ElapsedMilliseconds));
                EventLog.LogDebug(OverrideRoutingDomainDebugEnabled);

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in DomainReroutingAgent:OverrideRoutingDomain");
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
                int externalRecipients = 0;
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in DomainReroutingAgent:RemoveEmptyHeaders. OverrideRoutingDomainAgentEnabled is set to: {3}", messageId, sender, subject, OverrideRoutingDomainAgentEnabled));

                if (RemoveEmptyHeadersAgentEnabled && SenderOverrides.ContainsKey(sender))
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

                EventLog.AppendLogEntry(String.Format("DomainReroutingAgent:RemoveEmptyHeaders took {0} ms to execute", stopwatch.ElapsedMilliseconds));
                EventLog.LogDebug(RemoveEmptyHeadersDebugEnabled);

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in DomainReroutingAgent:RemoveEmptyHeaders");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }

    }

}