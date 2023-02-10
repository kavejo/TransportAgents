using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Exchange.Data.Transport.Smtp;
using Microsoft.Exchange.Data.Transport.Email;

namespace TransportAgents
{
    public class SenderDomainRewriteAgent : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new SenderSenderDomainRewriteAgent();
        }
    }

    public class SenderSenderDomainRewriteAgent : RoutingAgent
    {
        String oldDomainToDecommission = "contoso.com";
        String newDomainToUse = "tailspin.com";

        public SenderSenderDomainRewriteAgent()
        {
            base.OnSubmittedMessage += new SubmittedMessageEventHandler(RecipientDomainRewrite_OnSubmittedMessage);
        }

        void RecipientDomainRewrite_OnSubmittedMessage(SubmittedMessageEventSource source, QueuedMessageEventArgs evtMessage)
        {

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
                using (TextLogger textLog = new TextLogger(@"F:\\Transport Agents", @"RecipientDomainRewriteAgent_Log.txt"))
                {
                    StringBuilder errorEntry = new StringBuilder();

                    errorEntry.AppendLine("------------------------------------------------------------");
                    errorEntry.AppendLine("EXCEPTION!!!");
                    errorEntry.AppendLine("------------------------------------------------------------");
                    errorEntry.AppendLine(String.Format("HResult: {0}", ex.HResult.ToString()));
                    errorEntry.AppendLine(String.Format("Message: {0}", ex.Message.ToString()));
                    errorEntry.AppendLine(String.Format("Source: {0}", ex.Source.ToString()));
                    errorEntry.AppendLine(String.Format("InnerException: {0}", ex.InnerException.ToString()));
                    errorEntry.AppendLine(String.Format("StackTrace: {0}", ex.StackTrace.ToString()));

                    textLog.WriteToText(errorEntry.ToString(), "Error");
                }
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
                using (TextLogger textLog = new TextLogger(@"F:\\Transport Agents", @"SenderDomainRewriteAgent_Log.txt"))
                {
                    StringBuilder errorEntry = new StringBuilder();

                    errorEntry.AppendLine("------------------------------------------------------------");
                    errorEntry.AppendLine("EXCEPTION!!!");
                    errorEntry.AppendLine("------------------------------------------------------------");
                    errorEntry.AppendLine(String.Format("HResult: {0}", ex.HResult.ToString()));
                    errorEntry.AppendLine(String.Format("Message: {0}", ex.Message.ToString()));
                    errorEntry.AppendLine(String.Format("Source: {0}", ex.Source.ToString()));
                    errorEntry.AppendLine(String.Format("InnerException: {0}", ex.InnerException.ToString()));
                    errorEntry.AppendLine(String.Format("StackTrace: {0}", ex.StackTrace.ToString()));

                    textLog.WriteToText(errorEntry.ToString(), "Error");
                }
            }

            return;
        }
    }
}
