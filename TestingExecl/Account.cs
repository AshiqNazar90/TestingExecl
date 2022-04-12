using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestingExecl
{
    public class Account
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public long CreatedBy { get; set; }
        public long ModifiedBy { get; set; }

        public string EventName { get; set; }
        public string Location { get; set; }
        public double Load { get; set; }


    }
}
