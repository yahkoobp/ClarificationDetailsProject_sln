// Clarification.cs
// 
// This file contains the implementation of the Clarification class,
// which represents a clarification request in the application. It 
// encapsulates the properties of a clarification, such as its number, 
// document name, module, status, date, question, and answer.
// 
// Author: Yahkoob P
// Date: YYYY-MM-DD

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClarificationDetailsProject.Models
{
    /// <summary>
    /// Represents a clarification request with its associated properties.
    /// </summary>
    /// <remarks>
    /// This class is a data model that encapsulates the details of a 
    /// clarification, including its identifier, document name, module, 
    /// status, date, question, and answer. It is designed to be used 
    /// throughout the application to represent clarification requests.
    /// </remarks>
    public class Clarification
    {
        public int Number { get; set; }
        public string DocumentName { get; set; }
        public string Module { get; set; }
        public string Status { get; set; }
        public DateTime Date { get; set; }
        public string Question { get; set; }
        public string Answer { get; set; }
    }
}