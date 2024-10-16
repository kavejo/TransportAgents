﻿using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace TransportAgents
{

    public  class HeaderReroutingAgent : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new HeaderReroutingAgent_RoutingAgent(server.AcceptedDomains, server.AddressBook);
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

        static readonly string HeaderReroutingAgentName = "X-TransportAgent-Name";
        static readonly string HeaderReroutingAgentNameValue = "HeaderReroutingAgent";
        static readonly string HeaderReroutingAgentCreator = "X-TransportAgent-Creator";
        static readonly string HeaderReroutingAgentCreatorValue = "Tommaso Toniolo";
        static readonly string HeaderReroutingAgentContact = "X-TransportAgent-Contact";
        static readonly string HeaderReroutingAgentContactValue = "https://aka.ms/totoni";

        static readonly List<string> HeadersToRetain = new List<string>() 
        {
            "Accept-Language",
            "Bcc",
            "Cc",
            "Content-Language",
            "Content-Transfer-Encoding",
            "Content-Type",
            "Date",
            "From",
            "MIME-Version",
            "Message-ID",
            "Return-Path",
            "Subject",
            "Thread-Index",
            "Thread-Topic",
            "To",
            "X-MS-Exchange-Organization-Network-Message-Id",
            HeaderReroutingAgentTargetName,
            HeaderReroutingAgentP1P2MismatchActionName,
            HeaderReroutingAgentForceP1Name,
            HeaderReroutingAgentName,
            HeaderReroutingAgentCreator,
            HeaderReroutingAgentContact
        };
        
        static readonly string RegistryHive = @"Software\TransportAgents\HeaderReroutingAgent";
        static readonly string RegistryKeyDebugEnabledOverrideRoutingDomain = "DebugEnabled-OverrideRoutingDomain";
        static readonly string RegistryKeyDebugEnabledOverrideSenderAddress = "DebugEnabled-OverrideSenderAddress";
        static readonly string RegistryKeyDebugEnabledRemoveAllHeaders = "DebugEnabled-RemoveAllHeaders";
        static readonly string RegistryKeyDebugEnabled = "DebugEnabled";
        static bool DebugEnabledOverrideRoutingDomain = false;
        static bool DebugEnabledOverrideSenderAddress = false;
        static bool DebugEnabledRemoveAllHeaders = false;
        static bool DebugEnabled = false;
        static AcceptedDomainCollection acceptedDomains;
        static AddressBook addressBook;

        public HeaderReroutingAgent_RoutingAgent(AcceptedDomainCollection serverAcceptedDomains, AddressBook serverAddressBook)
        {
            base.OnResolvedMessage += new ResolvedMessageEventHandler(OverrideRoutingDomain);
            base.OnRoutedMessage += new RoutedMessageEventHandler(OverrideSenderAddress);
            base.OnCategorizedMessage += new CategorizedMessageEventHandler(RemoveUnnecessaryHeaders);

            RegistryKey registryPath = Registry.CurrentUser.OpenSubKey(RegistryHive, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
            if (registryPath != null)
            {
                string registryKeyValue = null;
                bool valueConversionResult = false;

                registryKeyValue = registryPath.GetValue(RegistryKeyDebugEnabledOverrideRoutingDomain, Boolean.FalseString).ToString();
                valueConversionResult = Boolean.TryParse(registryKeyValue, out DebugEnabledOverrideRoutingDomain);

                registryKeyValue = registryPath.GetValue(RegistryKeyDebugEnabledOverrideSenderAddress, Boolean.FalseString).ToString();
                valueConversionResult = Boolean.TryParse(registryKeyValue, out DebugEnabledOverrideSenderAddress);
                
                registryKeyValue = registryPath.GetValue(RegistryKeyDebugEnabledRemoveAllHeaders, Boolean.FalseString).ToString();
                valueConversionResult = Boolean.TryParse(registryKeyValue, out DebugEnabledRemoveAllHeaders);

                registryKeyValue = registryPath.GetValue(RegistryKeyDebugEnabled, Boolean.FalseString).ToString();
                valueConversionResult = Boolean.TryParse(registryKeyValue, out DebugEnabled);

                if (DebugEnabled == true) 
                {
                    DebugEnabledOverrideRoutingDomain = true;
                    DebugEnabledOverrideSenderAddress = true;
                    DebugEnabledRemoveAllHeaders = true;
                }

            }

            acceptedDomains = serverAcceptedDomains;
            addressBook = serverAddressBook;

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
                Header LoopPreventionHeader = headers.FindFirst(HeaderReroutingAgentName);

                if (HeaderReroutingAgentTarget != null && evtMessage.MailItem.Message.IsSystemMessage == false && LoopPreventionHeader == null)
                {
                    EventLog.AppendLogEntry(String.Format("Rerouting messages as the control header {0} is present", HeaderReroutingAgentTargetName));
                    HeaderReroutingAgentTargetValue = HeaderReroutingAgentTarget.Value.Trim();

                    if (!String.IsNullOrEmpty(HeaderReroutingAgentTargetValue) && (Uri.CheckHostName(HeaderReroutingAgentTargetValue) == UriHostNameType.Dns))
                    {
                        EventLog.AppendLogEntry(String.Format("Rerouting domain is valid as the header {0} is set to {1}", HeaderReroutingAgentTargetName, HeaderReroutingAgentTargetValue));

                        foreach (EnvelopeRecipient recipient in evtMessage.MailItem.Recipients)
                        {
                            EventLog.AppendLogEntry(String.Format("Evaluating recipient {0} which in Transport is categirized as {1}", recipient.Address.ToString(), recipient.RecipientCategory));
                            AcceptedDomain resolvedDomain = acceptedDomains.Find(recipient.Address.DomainPart.ToString());
                            EventLog.AppendLogEntry(String.Format("The check of whether the recipient domain is an Accepted Domain has returned {0}", resolvedDomain == null ? "NULL" : resolvedDomain.IsInCorporation.ToString()));
                            AddressBookEntry resolvedRecipient = addressBook.Find(recipient.Address);
                            EventLog.AppendLogEntry(String.Format("The check of whether the recipient in the Address Book has returned a type of {0}", resolvedRecipient == null ? "NULL" : resolvedRecipient.RecipientType.ToString()));
                            //bool isRecipientInternal = addressBook.IsInternal(recipient.Address);
                            //EventLog.AppendLogEntry(String.Format("The check of whether the recipient in Internal has returned {0}", isRecipientInternal));

                            if (resolvedDomain != null)
                            {
                                EventLog.AppendLogEntry(String.Format("Recipient {0} not overridden as the recipient domain IS AN ACCEPTED DOMAIN; the recipient is categorzed by Transport as {1}", recipient.Address.ToString(), recipient.RecipientCategory));
                            }
                            else
                            {
                                RoutingDomain customRoutingDomain = new RoutingDomain(HeaderReroutingAgentTargetValue);
                                RoutingOverride destinationOverride = new RoutingOverride(customRoutingDomain, DeliveryQueueDomain.UseOverrideDomain);
                                source.SetRoutingOverride(recipient, destinationOverride);
                                EventLog.AppendLogEntry(String.Format("Recipient {0} overridden to {1} as the recipient domain IS NOT AN ACCEPTED DOMAIN; the recipient is categorzed by Transport as {2}", recipient.Address.ToString(), HeaderReroutingAgentTargetValue, recipient.RecipientCategory));
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
                else 
                { 
                    if (evtMessage.MailItem.Message.IsSystemMessage == true) 
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as IsSystemMessage"));
                    }
                    else if (LoopPreventionHeader != null)
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is already present", LoopPreventionHeader.Name));
                        EventLog.AppendLogEntry(String.Format("This might mean there is a mail LOOP. Trace the message carefully."));
                        warningOccurred = true;
                    }
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is not set", HeaderReroutingAgentTargetName));
                    }
                }

                EventLog.AppendLogEntry(String.Format("HeaderReroutingAgent:OverrideRoutingDomain took {0} ms to execute", stopwatch.ElapsedMilliseconds));

                if (warningOccurred)
                {
                    EventLog.LogWarning();
                }
                else
                {
                    EventLog.LogDebug(DebugEnabledOverrideRoutingDomain);
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
                Header LoopPreventionHeader = headers.FindFirst(HeaderReroutingAgentName);

                if (HeaderReroutingAgentP1P2MismatchAction != null && evtMessage.MailItem.Message.IsSystemMessage == false && LoopPreventionHeader == null)
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
                else
                {
                    if (evtMessage.MailItem.Message.IsSystemMessage == true)
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as IsSystemMessage"));
                    }
                    else if (LoopPreventionHeader != null)
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is already present", LoopPreventionHeader.Name));
                        EventLog.AppendLogEntry(String.Format("This might mean there is a mail LOOP. Trace the message carefully."));
                        warningOccurred = true;
                    }
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is not set", HeaderReroutingAgentP1P2MismatchActionName));
                    }
                }

                Header HeaderReroutingAgentForceP1 = headers.FindFirst(HeaderReroutingAgentForceP1Name);

                if (HeaderReroutingAgentForceP1 != null && evtMessage.MailItem.Message.IsSystemMessage == false && LoopPreventionHeader == null)
                {
                    EventLog.AppendLogEntry(String.Format("Overriding P1 Sender as the control header {0} is present", HeaderReroutingAgentForceP1Name));
                    HeaderReroutingAgentForceP1Value = HeaderReroutingAgentForceP1.Value.Trim().ToUpper();

                    RoutingAddress newP1 = new RoutingAddress(HeaderReroutingAgentForceP1Value);
                    EventLog.AppendLogEntry(String.Format("The new P1 Sender will be forced is {0}", newP1.ToString()));

                    EventLog.AppendLogEntry(String.Format("P1 Sender is currently set to: {0}", P1Sender));
                    EventLog.AppendLogEntry(String.Format("P2 Sender is currently set to: {0}", P2Sender));

                    if (newP1.IsValid == false)
                    {
                        EventLog.AppendLogEntry(String.Format("The provided P1 Sender {0} IS INVALID", newP1.ToString()));
                        warningOccurred = true;
                    }
                    else
                    {
                        evtMessage.MailItem.FromAddress = newP1;
                        EventLog.AppendLogEntry(String.Format("Forced P1 Sender to {0}", evtMessage.MailItem.FromAddress.ToString()));
                    }
                }
                else
                {
                    if (evtMessage.MailItem.Message.IsSystemMessage == true)
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as IsSystemMessage"));
                    }
                    else if (LoopPreventionHeader != null)
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is already present", LoopPreventionHeader.Name));
                        EventLog.AppendLogEntry(String.Format("This might mean there is a mail LOOP. Trace the message carefully."));
                        warningOccurred = true;
                    }
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is not set", HeaderReroutingAgentForceP1Name));
                    }
                }

                EventLog.AppendLogEntry(String.Format("HeaderReroutingAgent:OverrideSenderAddress took {0} ms to execute", stopwatch.ElapsedMilliseconds));

                if (warningOccurred)
                {
                    EventLog.LogWarning();
                }
                else
                {
                    EventLog.LogDebug(DebugEnabledOverrideSenderAddress);
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

        void RemoveUnnecessaryHeaders(CategorizedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            try
            {
                bool warningOccurred = false; 
                string messageId = evtMessage.MailItem.Message.MessageId.ToString();
                string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();
                string subject = evtMessage.MailItem.Message.Subject.Trim();
                HeaderList headers = evtMessage.MailItem.Message.MimeDocument.RootPart.Headers;
                int externalRecipients = 0;
                Stopwatch stopwatch = Stopwatch.StartNew();

                EventLog.AppendLogEntry(String.Format("Processing message {0} from {1} with subject {2} in HeaderReroutingAgent:RemoveUnnecessaryHeaders", messageId, sender, subject));

                Header HeaderReroutingAgentTarget = headers.FindFirst(HeaderReroutingAgentTargetName);
                Header LoopPreventionHeader = headers.FindFirst(HeaderReroutingAgentName);

                if (HeaderReroutingAgentTarget != null && evtMessage.MailItem.Message.IsSystemMessage == false && LoopPreventionHeader == null)
                {
                    EventLog.AppendLogEntry(String.Format("Removing Unnecessary Headers as the control header {0} is present", HeaderReroutingAgentTargetName));

                    foreach (EnvelopeRecipient recipient in evtMessage.MailItem.Recipients)
                    {
                        EventLog.AppendLogEntry(String.Format("Evaluating recipient {0} which in Transport is categirized as {1}", recipient.Address.ToString(), recipient.RecipientCategory));
                        AcceptedDomain resolvedDomain = acceptedDomains.Find(recipient.Address.DomainPart.ToString());
                        EventLog.AppendLogEntry(String.Format("The check of whether the recipient domain is an Accepted Domain has returned {0}", resolvedDomain == null ? "NULL" : resolvedDomain.IsInCorporation.ToString()));
                        AddressBookEntry resolvedRecipient = addressBook.Find(recipient.Address);
                        EventLog.AppendLogEntry(String.Format("The check of whether the recipient in the Address Book has returned a type of {0}", resolvedRecipient == null ? "NULL" : resolvedRecipient.RecipientType.ToString()));
                        //bool isRecipientInternal = addressBook.IsInternal(recipient.Address);
                        //EventLog.AppendLogEntry(String.Format("The check of whether the recipient in Internal has returned {0}", isRecipientInternal));

                        if (resolvedDomain == null)
                        {
                            EventLog.AppendLogEntry(String.Format("The recipient {0} has been categorized as EXTERNAL as the recipient domain IS NOT AN ACCEPTED DOMAIN", recipient.Address.ToString().ToLower()));
                            externalRecipients++;
                        }
                    }

                    EventLog.AppendLogEntry(String.Format("There are {0} external recipient(s)", externalRecipients));
                    Dictionary<string, string> HeadersToInsert = new Dictionary<string, string>();
                    HeadersToInsert.Add(HeaderReroutingAgentName, HeaderReroutingAgentNameValue);
                    HeadersToInsert.Add(HeaderReroutingAgentCreator, HeaderReroutingAgentCreatorValue);
                    HeadersToInsert.Add(HeaderReroutingAgentContact, HeaderReroutingAgentContactValue);

                    if (externalRecipients > 0)
                    {
                        EventLog.AppendLogEntry(String.Format("Removing unnecessary or incompatible headers as there are external recipients"));
                        int MailHops = 1;

                        foreach (Header header in evtMessage.MailItem.Message.MimeDocument.RootPart.Headers)
                        {
                            if (HeadersToRetain.Contains(header.Name.ToLower(), StringComparer.OrdinalIgnoreCase))
                            {
                                EventLog.AppendLogEntry(String.Format("KEPT header {0}: {1}", header.Name, String.IsNullOrEmpty(header.Value) ? String.Empty : header.Value));
                            }
                            else
                            {
                                if (header.Name.ToLower() == "received")
                                {
                                    try
                                    {
                                        Regex RgxFrom = new Regex(@"(?<=from ).*(?= by)");
                                        Regex RgxBy = new Regex(@"(?<=by ).*(?=with)");
                                        HeadersToInsert.Add(String.Format("X-TransportAgent-ProcessHop-{0:000}", MailHops), String.Format("{0} by {1}", RgxFrom.Match(header.Value), RgxBy.Match(header.Value)));
                                        MailHops++;
                                    }
                                    catch 
                                    {
                                        EventLog.AppendLogEntry("Exception in Received-header Regex-parsing. Gracefully ignoring the same as it's a non-blocking one");
                                    }

                                }

                                if (header.Name.ToLower() == "x-ms-exchange-organization-rules-execution-history")
                                {
                                    HeadersToInsert.Add(String.Format("X-TransportAgent-ProcessRules"), header.Value);
                                }

                                evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.RemoveChild(header);
                                EventLog.AppendLogEntry(String.Format("REMOVED header {0}: {1}", header.Name, String.IsNullOrEmpty(header.Value) ? String.Empty : header.Value));
                            }
                        }
                    }

                    foreach (var newHeader in HeadersToInsert)
                    {
                        evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader(newHeader.Key, newHeader.Value), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                        EventLog.AppendLogEntry(String.Format("ADDED header {0}: {1}", newHeader.Key, String.IsNullOrEmpty(newHeader.Value) ? String.Empty : newHeader.Value));
                    }

                }
                else
                {
                    if (evtMessage.MailItem.Message.IsSystemMessage == true)
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as IsSystemMessage"));
                    }
                    else if (LoopPreventionHeader != null)
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is already present", LoopPreventionHeader.Name));
                        EventLog.AppendLogEntry(String.Format("This might mean there is a mail LOOP. Trace the message carefully."));
                        warningOccurred = true;
                    }
                    else
                    {
                        EventLog.AppendLogEntry(String.Format("Message has not been processed as {0} is not set", HeaderReroutingAgentTargetName));
                    }
                }

                EventLog.AppendLogEntry(String.Format("HeaderReroutingAgent:RemoveUnnecessaryHeaders took {0} ms to execute", stopwatch.ElapsedMilliseconds));

                if (warningOccurred)
                {
                    EventLog.LogWarning();
                }
                else
                {
                    EventLog.LogDebug(DebugEnabledRemoveAllHeaders);
                }

            }
            catch (Exception ex)
            {
                EventLog.AppendLogEntry("Exception in HeaderReroutingAgent:RemoveUnnecessaryHeaders");
                EventLog.AppendLogEntry(ex);
                EventLog.LogError();
            }

            return;

        }

    }
}
