using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace TransportAgents
{

    public class RecipientDomainRewriteAgent : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new RecipientDomainRewriteAgent_Agent();
        }
    }

    public class RecipientDomainRewriteAgent_Agent : RoutingAgent
    {

        static readonly string LogFile = String.Format("F:\\Transport Agents\\{0}.log", "RecipientDomainRewriteAgent");
        TextLogger TextLog = new TextLogger(LogFile);

        static readonly string oldDomainToDecommission = "contoso.com";
        static readonly string newDomainToUse = "tailspin.com";
        static readonly List<String> excludedFromRewrite = new List<String>(new String[] { "someone@testdomain.local", "someoneelse@testdomain.local"});


        public RecipientDomainRewriteAgent_Agent()
        {
            base.OnSubmittedMessage += new SubmittedMessageEventHandler(RecipientDomainRewrite_OnSubmittedMessage);
        }

        void RecipientDomainRewrite_OnSubmittedMessage(SubmittedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {

            TextLog.WriteToText("Entering: RecipientDomainRewrite_OnSubmittedMessage");
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
                for (int intCounter = evtMessage.MailItem.Recipients.Count - 1; intCounter >= 0; intCounter--)
                {

                    string msgRecipientP1 = evtMessage.MailItem.Recipients[intCounter].Address.LocalPart;
                    string msgDomainP1 = evtMessage.MailItem.Recipients[intCounter].Address.DomainPart;

                    if (msgDomainP1.ToLower() == oldDomainToDecommission.ToLower())
                    {
                        evtMessage.MailItem.Recipients[intCounter].Address = new RoutingAddress(msgRecipientP1, newDomainToUse);
                    }
                }
            }
            catch (Exception ex)
            {
                TextLog.WriteToText("------------------------------------------------------------");
                TextLog.WriteToText("EXCEPTION IN  P1 HEADER!!!");
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
                for (int intCounter = evtMessage.MailItem.Message.To.Count - 1; intCounter >= 0; intCounter--)
                {
                    int atIndex = evtMessage.MailItem.Message.To[intCounter].SmtpAddress.IndexOf("@");
                    int recLength = evtMessage.MailItem.Message.To[intCounter].SmtpAddress.Length;
                    string msgRecipientP2 = evtMessage.MailItem.Message.To[intCounter].SmtpAddress.Substring(0, atIndex);
                    string msgDomainP2 = evtMessage.MailItem.Message.To[intCounter].SmtpAddress.Substring(atIndex + 1, recLength - atIndex - 1);

                    if (msgDomainP2.ToLower() == oldDomainToDecommission.ToLower())
                    {
                        evtMessage.MailItem.Message.To[intCounter].SmtpAddress = msgRecipientP2 + "@" + newDomainToUse;
                    }
                }

                for (int intCounter = evtMessage.MailItem.Message.Cc.Count - 1; intCounter >= 0; intCounter--)
                {
                    int atIndex = evtMessage.MailItem.Message.Cc[intCounter].SmtpAddress.IndexOf("@");
                    int recLength = evtMessage.MailItem.Message.Cc[intCounter].SmtpAddress.Length;
                    string msgRecipientP2 = evtMessage.MailItem.Message.Cc[intCounter].SmtpAddress.Substring(0, atIndex);
                    string msgDomainP2 = evtMessage.MailItem.Message.Cc[intCounter].SmtpAddress.Substring(atIndex + 1, recLength - atIndex - 1);

                    if (msgDomainP2.ToLower() == oldDomainToDecommission.ToLower())
                    {
                        evtMessage.MailItem.Message.Cc[intCounter].SmtpAddress = msgRecipientP2 + "@" + newDomainToUse;
                    }
                }

                for (int intCounter = evtMessage.MailItem.Message.Bcc.Count - 1; intCounter >= 0; intCounter--)
                {
                    int atIndex = evtMessage.MailItem.Message.Bcc[intCounter].SmtpAddress.IndexOf("@");
                    int recLength = evtMessage.MailItem.Message.Bcc[intCounter].SmtpAddress.Length;
                    string msgRecipientP2 = evtMessage.MailItem.Message.Bcc[intCounter].SmtpAddress.Substring(0, atIndex);
                    string msgDomainP2 = evtMessage.MailItem.Message.Bcc[intCounter].SmtpAddress.Substring(atIndex + 1, recLength - atIndex - 1);

                    if (msgDomainP2.ToLower() == oldDomainToDecommission.ToLower())
                    {
                        evtMessage.MailItem.Message.Bcc[intCounter].SmtpAddress = msgRecipientP2 + "@" + newDomainToUse;
                    }
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
