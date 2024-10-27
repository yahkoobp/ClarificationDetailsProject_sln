// ----------------------------------------------------------------------------------------
// Project Name: ClarificationDetailsProject
// File Name: Clarification.cs
// Description: Defines a class for Clarification details
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
    /// Represents a clarification record containing information such as document name, module, status, date, question, and answer.
    /// </summary>
    public class Clarification
    {
        /// <summary>
        /// Gets or sets the unique identification number for the clarification.
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// Gets or sets the name of the document associated with this clarification.
        /// </summary>
        public string DocumentName { get; set; }

        /// <summary>
        /// Gets or sets the module associated with this clarification.
        /// </summary>
        public string Module { get; set; }

        /// <summary>
        /// Gets or sets the current status of this clarification (e.g., Pending, Closed).
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// Gets or sets the date of the clarification.
        /// </summary>
        public DateTime Date { get; set; }

        /// <summary>
        /// Gets or sets the question or issue being clarified.
        /// </summary>
        public string Question { get; set; }

        /// <summary>
        /// Gets or sets the answer or response to the clarification question.
        /// </summary>
        public string Answer { get; set; }
    }
}
