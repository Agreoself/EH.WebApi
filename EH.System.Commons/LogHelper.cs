using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Commons
{
    public class LogHelper
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public void LogInfo(string message)
        {
            logger.Info(message);
        }

        public void LogError(string message)
        {
            logger.Error(message);
        }

        public void LogWarn(string message)
        {
            logger.Warn(message);
        }

        public void LogDebug(string message)
        {
            logger.Debug(message);
        }

        public void LogFatal(string message)
        {
            logger.Fatal(message);
        }

        public void LogTrace(string message)
        {
            logger.Trace(message);
        }
    }
}
