using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Content
{
    [Serializable]
  public  class ExcelTypes
    {        
        public string value { get; set; }
        public string GroupId { get; set; }
        public string ParentId { get; set; }
        public List<string> Info { get; set; }        
        public List<KeyWords> keyWords { get; set; }
        public ExcelTypes()
        {                
        }
        public ExcelTypes(string _value, string _GoupId, string _ParentId, List<string> _info, List<KeyWords> _keys)
        {
            this.value = _value;
            this.GroupId = _GoupId;
            this.ParentId = _ParentId;
            this.Info = _info;
            this.keyWords = _keys;
        }
        public ExcelTypes(string _value, string _GoupId, string _ParentId)
        {
            this.value = _value;
            this.GroupId = _GoupId;
            this.ParentId = _ParentId;
        }
//        foreach (KeyValuePair<char, Person> keyValue in people)
//{
//    // keyValue.Value представляет класс Person
//    Console.WriteLine(keyValue.Key + " - " + keyValue.Value.Name); 
//}
 
//// перебор ключей
//foreach (char c in people.Keys)
//{
//    Console.WriteLine(c);
//}
 
//// перебор по значениям
//foreach (Person p in people.Values)
//{
//    Console.WriteLine(p.Name);
//}
    }
   
}
