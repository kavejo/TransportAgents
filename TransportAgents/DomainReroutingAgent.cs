using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
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
        TextLogger TextLog = new TextLogger(LogFile);

        public DomainReroutingAgent_RoutingAgent()
        {
            base.OnResolvedMessage += new ResolvedMessageEventHandler(DomainReroute_OnResolvedMessage);
        }

        void DomainReroute_OnResolvedMessage(ResolvedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {

            string messageId = evtMessage.MailItem.Message.MessageId.ToString();
            string sender = evtMessage.MailItem.FromAddress.ToString().ToLower().Trim();

            TextLog.WriteToText("Entering: DomainReroute_OnResolvedMessage");
            TextLog.WriteToText(String.Format("Processing message: {0} sent from <{1}>", messageId, sender));

            try
            {
                Stopwatch redirectionTime = new Stopwatch();
                redirectionTime.Start();

                if (sender == "noreply@toniolo.cloud")
                {
                    TextLog.WriteToText(String.Format("Rerouting the message as the sender is: {0}", sender));

                    foreach (var recipient in evtMessage.MailItem.Recipients)
                    {
                        var newRouteDomain = new RoutingDomain("acs.toniolo.cloud");
                        var dest = new RoutingOverride(newRouteDomain, DeliveryQueueDomain.UseOverrideDomain);
                        source.SetRoutingOverride(recipient, dest);
                        TextLog.WriteToText(String.Format("Routing domain overwritten for: {0}", recipient.Address));
                    }

                }

                redirectionTime.Stop();
                TextLog.WriteToText(String.Format("Message {0} processed in {1} ms", messageId, redirectionTime.Elapsed.Milliseconds));
            }
            catch (Exception ex)
            {
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText("EXCEPTION IN DOMAIN REDIRECTION!!!");
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText(String.Format("HResult: {0}", ex.HResult.ToString()));
                TextLog.WriteToText(String.Format("Message: {0}", ex.Message.ToString()));
                TextLog.WriteToText(String.Format("Source: {0}", ex.Source.ToString()));
            }

            return;

        }
    }
}
