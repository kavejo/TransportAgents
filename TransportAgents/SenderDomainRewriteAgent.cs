using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Collections.Generic;
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
        static readonly string LogFile = String.Format("F:\\Transport Agents\\{0}.log", "SenderDomainRewriteAgent");
        TextLogger TextLog = new TextLogger(LogFile);

        static readonly string oldDomainToDecommission = "contoso.com";
        static readonly string newDomainToUse = "tailspin.com";
        static readonly List<String> excludedFromRewrite = new List<String>(new String[] { "someone@testdomain.local", "someoneelse@testdomain.local" });

        public SenderDomainRewriteAgent_Agent()
        {
            base.OnSubmittedMessage += new SubmittedMessageEventHandler(SenderDomainRewrite_OnSubmittedMessage);
        }

        void SenderDomainRewrite_OnSubmittedMessage(SubmittedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {

            TextLog.WriteToText("Entering: SenderDomainRewrite_OnSubmittedMessage");
            TextLog.WriteToText(String.Format("{0}:<{1}><{2}><{3}>", "Processing message", evtMessage.MailItem.Message.MessageId.ToString(), evtMessage.MailItem.FromAddress.ToString().Trim(), evtMessage.MailItem.Message.Subject.ToString().Trim()));

            /////
            ///// As there is a need for not parsing the messages sent from specific senders (listed in List<String> excludedFromRewrite).
            ///// In this case the class will return to the caller before executing any of the rewrite-logic.
            /////
            try
            {
                if (excludedFromRewrite.Contains(evtMessage.MailItem.FromAddress.ToString().Trim()) || excludedFromRewrite.Contains(evtMessage.MailItem.Message.Sender.ToString().Trim()))
                {
                    TextLog.WriteToText(String.Format("{0}:P1={1}/P2={2}", "Avoid to process as the sender is", evtMessage.MailItem.FromAddress.ToString().Trim(), evtMessage.MailItem.Message.Sender.ToString().Trim()));
                    return;
                }
            }
            catch (Exception ex)
            {
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText("EXCEPTION IN SENDER CHECK!!!");
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText(String.Format("HResult: {0}", ex.HResult.ToString()));
                TextLog.WriteToText(String.Format("Message: {0}", ex.Message.ToString()));
                TextLog.WriteToText(String.Format("Source: {0}", ex.Source.ToString()));
            }

            /////
            ///// P1 HEADER (SMTP envelope)
            /////
            try
            {
                string msgSenderP1 = evtMessage.MailItem.FromAddress.LocalPart;
                string msgDomainP1 = evtMessage.MailItem.FromAddress.DomainPart;

                if (msgDomainP1.ToLower() == oldDomainToDecommission.ToLower())
                {
                    evtMessage.MailItem.FromAddress = new RoutingAddress(msgSenderP1, newDomainToUse);
                }
            }
            catch (Exception ex)
            {
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText("EXCEPTION IN P1 HEADER!!!");
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText(String.Format("HResult: {0}", ex.HResult.ToString()));
                TextLog.WriteToText(String.Format("Message: {0}", ex.Message.ToString()));
                TextLog.WriteToText(String.Format("Source: {0}", ex.Source.ToString()));
            }

            /////
            ///// P2 HEADER (Message envelope)
            /////
            try
            {
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
                TextLog.WriteToText("EXCEPTION IN P2 HEADER!!!");
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText(String.Format("HResult: {0}", ex.HResult.ToString()));
                TextLog.WriteToText(String.Format("Message: {0}", ex.Message.ToString()));
                TextLog.WriteToText(String.Format("Source: {0}", ex.Source.ToString()));
            }

            return;

        }
    }
}
