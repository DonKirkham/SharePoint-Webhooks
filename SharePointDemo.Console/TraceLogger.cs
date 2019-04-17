using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace SharePointDemo.Console
{
    //public interface ILogger
    //{
    //    void Error(string message);
    //    void Information(string message);
    //    void Warning(string message);
    //}
    /// <summary>
    /// Windows Trace logger
    /// </summary>
    public class TraceLogger : ILogger
    {
        public TraceLogger()
        {
            System.Diagnostics.Trace.Listeners.Add(new System.Diagnostics.TextWriterTraceListener(System.Console.Out));
        }

        public IDisposable BeginScope<TState>(TState state)
        {
            return null;
        }

        public bool IsEnabled(LogLevel logLevel)
        {
            return true;
        }

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception exception, Func<TState, Exception, string> formatter)
        {
            string message = string.Empty;

            if (formatter != null)
            {
                message = formatter(state, exception);
            }
            else
            {
                //message = LogFormatter.Formatter(state, exception);
            }

            System.Console.WriteLine(message);
        }
    }
}
