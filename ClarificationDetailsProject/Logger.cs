﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;

namespace ClarificationDetailsProject
{
    public static class Logger
    {
        public static readonly ILog log = LogManager.GetLogger(typeof(Logger));
    }
}
