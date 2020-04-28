using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nndk
{
    public class Record
    {
        public string code { get; set; }
        public string name { get; set; }
        public DateTime? time { get; set; }
    }
    public class Result
    {
        public string code { get; set; }
        public string name { get; set; }
        public double sum { get; set; }
    }
}
