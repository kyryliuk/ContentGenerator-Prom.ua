using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Content
{
  public  class ExcelTypes
    {
        public string value { get; set; }
        public string GroupId { get; set; }
        public string ParentId { get; set; }
        public List<string> Info { get; set; }
        public List<List<string>> Keys { get; set; }
        public ExcelTypes()
        {
                
        }
        public ExcelTypes(string _value, string _GoupId, string _ParentId, List<string> _info, List<List<string>> _keys)
        {
            this.value = _value;
            this.GroupId = _GoupId;
            this.ParentId = _ParentId;
            this.Info = _info;
            this.Keys = _keys;
        }
        public ExcelTypes(string _value, string _GoupId, string _ParentId)
        {
            this.value = _value;
            this.GroupId = _GoupId;
            this.ParentId = _ParentId;
        }
    }
   
}
