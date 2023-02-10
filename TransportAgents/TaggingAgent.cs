using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Exchange.Data.Transport.Smtp;

namespace TransportAgents
{
    public class TaggingAgent : SmtpReceiveAgentFactory
    {
        public override SmtpReceiveAgent CreateAgent(SmtpServer server)
        {
            return new TaggingAgent_Agent();
        }
    }

    public partial class TaggingAgent_Agent : SmtpReceiveAgent
    {

        public TaggingAgent_Agent()
        {
            OnRcptCommand += RcptMessageHandler;
        }

        private void RcptMessageHandler(ReceiveCommandEventSource receiveMessageEventSource, RcptCommandEventArgs eventArgs)
        {
            RoutingAddress initialRecipientAddress = eventArgs.RecipientAddress;
            try
            {
                int plusIndex = initialRecipientAddress.LocalPart.IndexOf("+");
                string recipient = initialRecipientAddress.LocalPart.Substring(0, plusIndex);
                string tag = initialRecipientAddress.LocalPart.Substring(plusIndex + 1);
                string domain = initialRecipientAddress.DomainPart;
                string revisedRecipientAddress = recipient + "@" + domain;
                eventArgs.RecipientAddress = RoutingAddress.Parse(revisedRecipientAddress);
                eventArgs.OriginalRecipient = revisedRecipientAddress;
            }
            catch (Exception ex)
            {
                using (TextLogger textLog = new TextLogger(@"F:\\Transport Agents", @"TaggingAgent_Log.txt"))
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
        }
    }

}

