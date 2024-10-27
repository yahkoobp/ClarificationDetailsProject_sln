// ----------------------------------------------------------------------------------------
// Project Name: ClarificationDetailsProject
// File Name: Summary.cs
// Description: Defines a class for summary
// Author: Yahkoob P
// Date: 27-10-2024
// ----------------------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClarificationDetailsProject.Models
{
    /// <summary>
    /// Represents a summary of clarifications for a particular module, including counts of various statuses.
    /// </summary>
    public class Summary
    {
        /// <summary>
        /// Gets or sets the unique identifier for the summary.
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// Gets or sets the name of the module associated with this summary.
        /// </summary>
        public string Module { get; set; }

        /// <summary>
        /// Gets or sets the count of closed clarifications.
        /// </summary>
        public int Closed { get; set; }

        /// <summary>
        /// Gets or sets the count of open clarifications.
        /// </summary>
        public int Open { get; set; }

        /// <summary>
        /// Gets or sets the count of clarifications that are on hold.
        /// </summary>
        public int OnHold { get; set; }

        /// <summary>
        /// Gets or sets the count of pending clarifications.
        /// </summary>
        public int Pending { get; set; }

        /// <summary>
        /// Gets or sets the total count of clarifications across all statuses.
        /// </summary>
        public int Total { get; set; }
    }
}
