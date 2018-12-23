using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Content
{
  public  class twolist
    {
        public string ListValues { get; set; }
        public string ListValues1 { get; set; }
        public twolist()
        {

        }
        public twolist(string _ListValues, string _ListValues1)
        {
            this.ListValues = _ListValues;
            this.ListValues1 = _ListValues1;
        }
    }
}
