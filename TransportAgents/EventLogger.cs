using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransportAgents
{
    class EventLogger
    {

        public static void WriteToEventLog(string message, EventLogEntryType severity)
        {
            using (EventLog eventLog = new EventLog("Application"))
            {
                eventLog.Source = "TransportAgents";
                eventLog.WriteEntry(message, severity);
            }
        }

        public static void WriteToEventLog(string message)
        {
            WriteToEventLog(message, EventLogEntryType.Information);
        }
    }
}
