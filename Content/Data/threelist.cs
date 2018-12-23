using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Content
{  
    public  class threelist
    {
        public string ListValues1 { get; set; }
        public string ListValues2 { get; set; }
        public string ListValues3 { get; set; }
        public threelist()
        {

        }
        public threelist(string _ListValues1, string _ListValues2, string _ListValues3)
        {
            this.ListValues1 = _ListValues1;
            this.ListValues2 = _ListValues2;
            this.ListValues3 = _ListValues3;
        }
    }
}
