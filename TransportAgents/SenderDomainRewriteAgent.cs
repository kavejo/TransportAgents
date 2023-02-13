using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Diagnostics;
using System.Text;

namespace TransportAgents
{
    public class SenderDomainRewriteAgent : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new SenderDomainRewriteAgent_Agent();
        }
    }

    public class SenderDomainRewriteAgent_Agent : RoutingAgent
    {
        static string LogFile = String.Format("F:\\Transport Agents\\{0}.log", "SenderDomainRewriteAgent");
        TextLogger TextLog = new TextLogger(LogFile);

        static string oldDomainToDecommission = "contoso.com";
        static string newDomainToUse = "tailspin.com";

        public SenderDomainRewriteAgent_Agent()
        {
            base.OnSubmittedMessage += new SubmittedMessageEventHandler(SenderDomainRewrite_OnSubmittedMessage);
        }

        void SenderDomainRewrite_OnSubmittedMessage(SubmittedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {

            TextLog.WriteToText("Entering: SenderDomainRewrite_OnSubmittedMessage");

            try
            {

                /////
                ///// P1 HEADER (SMTP envelope)
                /////

                string msgSenderP1 = evtMessage.MailItem.FromAddress.LocalPart;
                string msgDomainP1 = evtMessage.MailItem.FromAddress.DomainPart;

                if (msgDomainP1.ToLower() == oldDomainToDecommission.ToLower())
                {
                    evtMessage.MailItem.FromAddress = new RoutingAddress(msgSenderP1, newDomainToUse);
                }

                /////
                ///// P2 HEADER (Message envelope)
                /////

                int atIndex = evtMessage.MailItem.Message.From.SmtpAddress.IndexOf("@");
                int recLength = evtMessage.MailItem.Message.From.SmtpAddress.Length;
                string msgSenderP2 = evtMessage.MailItem.Message.From.SmtpAddress.Substring(0, atIndex);
                string msgDomainP2 = evtMessage.MailItem.Message.From.SmtpAddress.Substring(atIndex + 1, recLength - atIndex - 1);

                if (msgDomainP2.ToLower() == oldDomainToDecommission.ToLower())
                {
                    evtMessage.MailItem.Message.From.SmtpAddress = msgSenderP2 + "@" + newDomainToUse;
                    evtMessage.MailItem.Message.Sender.SmtpAddress = msgSenderP2 + "@" + newDomainToUse;
                }
            }
            catch (Exception ex)
            {
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText("EXCEPTION!!!");
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText(String.Format("HResult: {0}", ex.HResult.ToString()));
                TextLog.WriteToText(String.Format("Message: {0}", ex.Message.ToString()));
                TextLog.WriteToText(String.Format("Source: {0}", ex.Source.ToString()));
            }

            return;

        }
    }
}
