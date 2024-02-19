using System;
using System.Diagnostics;
using System.Diagnostics.Tracing;
using System.Linq;
using System.Reflection;
using System.Security;

namespace TransportAgents
{
    internal class EventLogger
    {
        private string Source = String.Empty;

        public EventLogger(string eventSource) 
        {
            //try
            //{
            bool sourceExists = EventLog.SourceExists(eventSource);
            Source = eventSource;
            if (sourceExists == false)
            {
                EventLog.CreateEventSource(Source, "Application");
            }
            //}
            //catch (Exception ex)
            //{
            //    Source = "Application";
            //    if (!EventLog.SourceExists(Source))
            //    {
            //        EventLog.CreateEventSource(Source, "Application");
            //    }
            //}
        }

        public void LogDebug(string message, bool isDebugEnabled)
        {
            if (isDebugEnabled)
            {
                EventLog.WriteEntry(Source, message, EventLogEntryType.Information);
            }
        }

        public void LogInformation(string message)
        {
            EventLog.WriteEntry(Source, message, EventLogEntryType.Information);
        }

        public void LogWarning(string message)
        {
            EventLog.WriteEntry(Source, message, EventLogEntryType.Warning);
        }

        public void LogError(string message)
        {
            EventLog.WriteEntry(Source, message, EventLogEntryType.Error);
        }

        public void LogException(Exception ex)
        {
            EventLog.WriteEntry(Source, ex.ToString(), EventLogEntryType.Error);
        }

    }
}
