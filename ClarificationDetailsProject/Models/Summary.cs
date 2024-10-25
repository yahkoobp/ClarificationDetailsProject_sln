using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClarificationDetailsProject.Models
{
    public class Summary
    {
        public int Number {  get; set; }
        public string Module {  get; set; }
        public int Closed { get; set; }
        public int Open {  get; set; }
        public int OnHold {  get; set; }
        public int Pending {  get; set; }
        public int Total {  get; set; }
    }
}
