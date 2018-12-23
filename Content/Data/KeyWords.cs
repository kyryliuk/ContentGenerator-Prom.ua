using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Content
{
    [Serializable]
    public class KeyWords
    {
        public string KeyName { get; set; }
        public List<string> ListValues{ get; set; }
        public List<string> ListValues1 { get; set; }
        public KeyWords()
        {
               
        }
        public KeyWords(string _KeyName, List<string> _ListValues, List<string> _ListValues1)
        {
            this.KeyName = _KeyName;
            this.ListValues = _ListValues;
            this.ListValues1 = _ListValues1;
        }
      
    }
    
}
