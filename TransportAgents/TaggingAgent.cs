using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Smtp;
using System;
using System.Diagnostics;
using System.Text;

namespace TransportAgents
{
    public class TaggingAgent : SmtpReceiveAgentFactory
    {
        public override SmtpReceiveAgent CreateAgent(SmtpServer server)
        {
            return new TaggingAgent_Agent();
        }
    }

    public class TaggingAgent_Agent : SmtpReceiveAgent
    {

        static string LogFile = String.Format("F:\\Transport Agents\\{0}.log", "TaggingAgent");
        TextLogger TextLog = new TextLogger(LogFile);

        public TaggingAgent_Agent()
        {
            OnRcptCommand += StripTagOutOfAddress;
        }
    
        private void StripTagOutOfAddress(ReceiveCommandEventSource receiveMessageEventSource, RcptCommandEventArgs eventArgs)
        {
            TextLog.WriteToText("Entering: StripTagOutOfAddress");

            RoutingAddress initialRecipientAddress = eventArgs.RecipientAddress;
            try
            {
                if (initialRecipientAddress.LocalPart.Contains("+"))
                {
                    int plusIndex = initialRecipientAddress.LocalPart.IndexOf("+");
                    string recipient = initialRecipientAddress.LocalPart.Substring(0, plusIndex);
                    string tag = initialRecipientAddress.LocalPart.Substring(plusIndex + 1);
                    string domain = initialRecipientAddress.DomainPart;
                    string revisedRecipientAddress = recipient + "@" + domain;
                    eventArgs.RecipientAddress = RoutingAddress.Parse(revisedRecipientAddress);
                    eventArgs.OriginalRecipient = revisedRecipientAddress;
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

