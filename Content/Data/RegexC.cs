using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Content
{
    [Serializable]
    public   class RegexC
    {
        public string Regex_func { get; set; }
        public bool check { get; set; }
        public string Value { get; set; }

            public RegexC(string _Regex_func, string _value, bool _check )
            {
            this.Regex_func = _Regex_func;
            this.Value = _value;
            this.check = _check;
            }

            public RegexC()
            {

            }
    }
   
}
