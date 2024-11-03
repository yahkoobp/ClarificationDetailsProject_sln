// ----------------------------------------------------------------------------------------
// Project Name: ClarificationDetailsProject
// File Name: Logger.cs
// Description: Defines a static class for logger
// Author: Yahkoob P
// Date: 27-10-2024
// ----------------------------------------------------------------------------------------

using log4net;

namespace ClarificationDetailsProject
{
    /// <summary>
    /// Defines a static class for logger
    /// </summary>
    public static class Logger
    {
        /// <summary>
        /// creates an object for the logger class
        /// </summary>
        public static readonly ILog log = LogManager.GetLogger(typeof(Logger));
    }
}
