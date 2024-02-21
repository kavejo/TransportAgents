using System;
using System.Diagnostics;
using System.Security;
using System.Text;

namespace TransportAgents
{
    internal class EventLogger : IDisposable
    {
        private string Source = String.Empty;
        private StringBuilder EventLogMessage = null;

        public EventLogger(string eventSource = "Application") 
        {
            EventLogMessage = new StringBuilder();

            try
            {
                bool sourceExists = EventLog.SourceExists(eventSource);
                Source = eventSource;
                if (sourceExists == false)
                {
                    EventLog.CreateEventSource(Source, "Application");
                }
            }
            catch (SecurityException)
            {
                Source = "Application";
                if (!EventLog.SourceExists(Source))
                {
                    EventLog.CreateEventSource(Source, "Application");
                }
            }
        }

        public void LogDebug(string message, bool isDebugEnabled = true)
        {
            if (isDebugEnabled)
            {
                EventLog.WriteEntry(Source, message, EventLogEntryType.Information);
            }
        }

        public void LogDebug(bool isDebugEnabled = true)
        {
            if (isDebugEnabled)
            {
                EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Information);
                EventLogMessage.Clear();
            }
        }

        public void LogInformation(string message)
        {
            EventLog.WriteEntry(Source, message, EventLogEntryType.Information);
        }

        public void LogInformation()
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Information);
            EventLogMessage.Clear();
        }

        public void LogWarning(string message)
        {
            EventLog.WriteEntry(Source, message, EventLogEntryType.Warning);
        }

        public void LogWarning()
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Warning);
            EventLogMessage.Clear();
        }

        public void LogError(string message)
        {
            EventLog.WriteEntry(Source, message, EventLogEntryType.Error);
        }

        public void LogError()
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Error);
            EventLogMessage.Clear();
        }

        public void LogException(Exception ex)
        {
            EventLog.WriteEntry(Source, ex.ToString(), EventLogEntryType.Error);
        }

        public void LogException()
        {
            EventLog.WriteEntry(Source, EventLogMessage.ToString(), EventLogEntryType.Error);
            EventLogMessage.Clear();
        }

        public void AppendLogEntry(string message)
        {
            EventLogMessage.AppendLine(message);
        }

        public void AppendLogEntry(Exception ex)
        {
            EventLogMessage.AppendLine("--------------------------------------------------------------------------------");
            EventLogMessage.AppendLine(String.Format("EXCEPTION MESSAGE: {0}", ex.Message));
            EventLogMessage.AppendLine(String.Format("EXCEPTION HRESULT: {0}", ex.HResult));
            EventLogMessage.AppendLine(String.Format("EXCEPTION SOURCE: {0}", ex.Source));
            EventLogMessage.AppendLine(String.Format("EXCEPTION INNER EXCEPTION: {0}", ex.InnerException));
            EventLogMessage.AppendLine(String.Format("EXCEPTION STRACK: {0}", ex.StackTrace));
            EventLogMessage.AppendLine("--------------------------------------------------------------------------------");
            EventLogMessage.AppendLine(ex.ToString());
            EventLogMessage.AppendLine("--------------------------------------------------------------------------------");
        }

        public void AppendLogEntry(object obj)
        {
            EventLogMessage.AppendLine(obj.ToString());
        }

        public void ClearLogEntry()
        {
            EventLogMessage.Clear();
        }

        public string GetLogEntry()
        {
            return EventLogMessage.ToString();
        }

        ~EventLogger()
        {
            WriteEventLogOnExit();
        }

        void IDisposable.Dispose()
        {
            WriteEventLogOnExit();
        }

        private void WriteEventLogOnExit()
        {
            if (!String.IsNullOrEmpty(EventLogMessage.ToString()))
            {
                EventLogMessage.AppendLine("Writing Event on Agent exit");
                LogInformation(EventLogMessage.ToString());
                EventLogMessage.Clear();
            }
        }
    }
}
