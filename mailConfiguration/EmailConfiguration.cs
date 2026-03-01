using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StatementFile.StatementFile.mailConfiguration
{
    public class To
    {
        public string email { get; set; }
        public string name { get; set; }
        public bool? valid { get; set; }
    }

    public class Cc
    {
        public string email { get; set; }
        public string name { get; set; }
    }

    
    public class EmailConfiguration
    {
        public string name { get; set; }
        public string message { get; set; }
        public List<To> to { get; set; }
        public List<Cc> cc { get; set; }
        public string path { get; set; }
    }

    
}
