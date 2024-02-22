using System;
using System.IO;

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
            catch (Exception)
            {
            }
        }

        public TextLogger(string logPath)
        {
            if (_logPath != logPath)
            {
                CloseStream();
            }

            try
            {
                _logPath = logPath;
                _logStream = new StreamWriter(_logPath, true);
            }
            catch (Exception)
            {
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

        public void WriteToText(string message)
        {
            _logStream.WriteLine(String.Format("{0:yyyy-MM-dd HH:mm:ss} | {1}", DateTime.Now, message));
            _logStream.Flush();
        }

    }
}
