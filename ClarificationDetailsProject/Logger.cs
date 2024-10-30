using log4net;

namespace ClarificationDetailsProject
{
    public static class Logger
    {
        public static readonly ILog log = LogManager.GetLogger(typeof(Logger));
    }
}
