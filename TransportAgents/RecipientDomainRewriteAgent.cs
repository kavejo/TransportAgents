using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using System;
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

        static string LogFile = String.Format("F:\\Transport Agents\\{0}.log", "RecipientDomainRewriteAgent");
        TextLogger TextLog = new TextLogger(LogFile);

        static string oldDomainToDecommission = "contoso.com";
        static string newDomainToUse = "tailspin.com";

        public RecipientDomainRewriteAgent_Agent()
        {
            base.OnSubmittedMessage += new SubmittedMessageEventHandler(RecipientDomainRewrite_OnSubmittedMessage);
        }

        void RecipientDomainRewrite_OnSubmittedMessage(SubmittedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {

            TextLog.WriteToText("Entering: RecipientDomainRewrite_OnSubmittedMessage");

            try
            {
                /////
                ///// P1 HEADER (SMTP envelope)
                /////

                for (int intCounter = evtMessage.MailItem.Recipients.Count - 1; intCounter >= 0; intCounter--)
                {

                    string msgRecipientP1 = evtMessage.MailItem.Recipients[intCounter].Address.LocalPart;
                    string msgDomainP1 = evtMessage.MailItem.Recipients[intCounter].Address.DomainPart;

                    if (msgDomainP1.ToLower() == oldDomainToDecommission.ToLower())
                    {
                        evtMessage.MailItem.Recipients[intCounter].Address = new RoutingAddress(msgRecipientP1, newDomainToUse);
                    }
                }

                /////
                ///// P2 HEADER (Message envelope)
                /////

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
