using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
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
        static readonly string LogFile = String.Format("F:\\Transport Agents\\{0}.log", "DomainReroutingAgent");
        static readonly List<string> SendersToReroute = new List<string> { "noreply@toniolo.cloud" };
        static readonly RoutingDomain customRoutingDomain = new RoutingDomain("acs.toniolo.cloud");
        static readonly RoutingOverride destinationOverride = new RoutingOverride(customRoutingDomain, DeliveryQueueDomain.UseOverrideDomain);
        TextLogger TextLog = new TextLogger(LogFile);

        public DomainReroutingAgent_RoutingAgent()
        {
            base.OnResolvedMessage += new ResolvedMessageEventHandler(SendViaCustomRoutingDomain);
            base.OnCategorizedMessage += new CategorizedMessageEventHandler(RemoveUnsupportedHeaders);
        }

        void SendViaCustomRoutingDomain(ResolvedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            string messageId = evtMessage.MailItem.Message.MessageId.ToString();
            string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();

            TextLog.WriteToText("----------------------------------------------------------------------------------------------------");
            TextLog.WriteToText(String.Format("SendViaCustomRoutingDomain - Processing message: {0} sent from <{1}>", messageId, sender));
            TextLog.WriteToText("----------------------------------------------------------------------------------------------------");

            try
            {
                Stopwatch redirectionTime = new Stopwatch();
                redirectionTime.Start();

                if ( SendersToReroute.Contains(sender) )
                {
                    TextLog.WriteToText(String.Format("Rerouting the message as the sender is: {0}", sender));

                    foreach (EnvelopeRecipient recipient in evtMessage.MailItem.Recipients)
                    {
                        if (recipient.RecipientCategory == RecipientCategory.InDifferentOrganization || recipient.RecipientCategory == RecipientCategory.Unknown)
                        {
                            source.SetRoutingOverride(recipient, destinationOverride);
                            TextLog.WriteToText(String.Format("Routing domain overwritten for: {0} as the recipient is external", recipient.Address));
                        }
                        else
                        {
                            TextLog.WriteToText(String.Format("Routing domain not overwritten for: {0} as the recipient is internal", recipient.Address));
                        }
                    }

                    //foreach (Header header in evtMessage.MailItem.Message.MimeDocument.RootPart.Headers)
                    //{
                    //    TextLog.WriteToText(String.Format("Processing: {0}: {1}", header.Name, header.Value));
                    //    if (header.Value == null || header.Value.Length == 0 || String.IsNullOrEmpty(header.Value))
                    //    {
                    //        evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.RemoveChild(header);
                    //        TextLog.WriteToText(String.Format("Removed as EMPTY"));
                    //    }
                    //    if (header.Name == "Received")
                    //    {
                    //        evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.RemoveChild(header);
                    //        TextLog.WriteToText(String.Format("Removed as of type RECEIVED"));
                    //    }
                    //}

                    //TextLog.WriteToText("Adding custom headers");
                    //evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Name", "DomainReroutingAgent"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                    //evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Creator", "Tommaso Toniolo"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                    //evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Contact", "https://aka.ms/totoni"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);

                }

                redirectionTime.Stop();
                TextLog.WriteToText(String.Format("SendViaCustomRoutingDomain executed for message {0} in {1} ms", messageId, redirectionTime.Elapsed.Milliseconds));
            }
            catch (Exception ex)
            {
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText("EXCEPTION IN SENDVIACUSTOMROUTINGDOMAIN!!!");
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText(String.Format("HResult: {0}", ex.HResult.ToString()));
                TextLog.WriteToText(String.Format("Message: {0}", ex.Message.ToString()));
                TextLog.WriteToText(String.Format("Source: {0}", ex.Source.ToString()));
            }

            return;

        }

        void RemoveUnsupportedHeaders(CategorizedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {
            string messageId = evtMessage.MailItem.Message.MessageId.ToString();
            string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();

            TextLog.WriteToText("----------------------------------------------------------------------------------------------------");
            TextLog.WriteToText(String.Format("RemoveUnsupportedHeaders - Processing message: {0} sent from <{1}>", messageId, sender));
            TextLog.WriteToText("----------------------------------------------------------------------------------------------------");
            
            try
            {
                Stopwatch headerRemovalTime = new Stopwatch();
                headerRemovalTime.Start();

                if (SendersToReroute.Contains(sender))
                {
                    TextLog.WriteToText(String.Format("Removing unsupported headers the message as the sender is: {0}", sender));

                    foreach (Header header in evtMessage.MailItem.Message.MimeDocument.RootPart.Headers)
                    {
                        TextLog.WriteToText(String.Format("Processing: {0}: {1}", header.Name, header.Value));
                        if (header.Value == null || header.Value.Length == 0 || String.IsNullOrEmpty(header.Value))
                        {
                            if (header.Name != "From" && header.Name != "To")
                            {
                                evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.RemoveChild(header);
                                TextLog.WriteToText(String.Format("Removed as EMPTY"));
                            }
                            else
                            {
                                evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.RemoveChild(header);
                                TextLog.WriteToText(String.Format("Kept EMPTY {0} as it will be updated later", header.Name));
                            }
                        }
                        if (header.Name == "Received")
                        {
                            evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.RemoveChild(header);
                            TextLog.WriteToText(String.Format("Removed as of type RECEIVED"));
                        }
                    }

                    TextLog.WriteToText("Adding custom headers");
                    evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Name", "DomainReroutingAgent"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                    evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Creator", "Tommaso Toniolo"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);
                    evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.InsertAfter(new TextHeader("X-TransportAgent-Contact", "https://aka.ms/totoni"), evtMessage.MailItem.Message.MimeDocument.RootPart.Headers.LastChild);

                }

                headerRemovalTime.Stop();
                TextLog.WriteToText(String.Format("RemoveUnsupportedHeaders executed for message {0} in {1} ms", messageId, headerRemovalTime.Elapsed.Milliseconds));
            }
            catch (Exception ex)
            {
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText("EXCEPTION IN REMOVEUNSUPPORTEDHEADERS!!!");
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText(String.Format("HResult: {0}", ex.HResult.ToString()));
                TextLog.WriteToText(String.Format("Message: {0}", ex.Message.ToString()));
                TextLog.WriteToText(String.Format("Source: {0}", ex.Source.ToString()));
            }

            return;

        }

    }

}
