using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace TransportAgents
{
    public class TextLogger : IDisposable
    {
        private string _logPath = string.Empty;
        private StreamWriter _logStream = null;

        public TextLogger(string logLocation, string logName)
        {
            string logPath = Path.Combine(logLocation, logName);

            if (_logPath != logPath)
            {
                CloseStream();
            }

            try
            {
                _logPath = logPath;
                _logStream = new StreamWriter(_logPath, true);
            }
            catch (Exception e)
            {
                EventLogger.WriteToEventLog(e.Message, EventLogEntryType.Error);
            }
        }

        ~TextLogger()
        {
            CloseStream();
        }

        void IDisposable.Dispose()
        {
            CloseStream();
        }
        private void CloseStream()
        {
            if (_logStream != null)
            {
                _logStream.Flush();
                _logStream.Dispose();
            }
        }

        public void WriteToText(string message, string severity)
        {

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
            sb.AppendLine(String.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Now));
            sb.AppendLine(String.Format("Severity: {0}", severity));
            sb.AppendLine(message);
            sb.AppendLine("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
            _logStream.WriteLine(sb.ToString());
            _logStream.Flush();
        }

        public void WriteToText(string message)
        {
            WriteToText(message, "Information");
        }
    }
}
