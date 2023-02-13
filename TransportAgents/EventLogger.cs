using System.Diagnostics;

namespace TransportAgents
{
    class EventLogger
    {

        public static void WriteToEventLog(string source, EventLogEntryType severity, string message)
        {
            using (EventLog eventLog = new EventLog("Application"))
            {
                eventLog.Source = source;
                eventLog.WriteEntry(message, severity);
            }
        }

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
