using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
/// <summary>
/// клас для генерування ключових слів саме в ексель там де є колонка ключові слова
/// </summary>
namespace Content
{
    [Serializable]
   public class MainWords
    {
        public string GroupId { get; set; }
        public string ParentId { get; set; }
        public List<string> Words { get; set; }
        public string value { get; set; }
        public MainWords()
        {
                
        }
        public MainWords(string _GroupId, string _ParentId, List<string> _words, string _value)
        {
            this.GroupId = _GroupId;
            this.ParentId = _ParentId;
            this.Words = _words;
            this.value = _value;
        }
        public MainWords(string _value, string _GroupId, string _ParentId)
        {
            this.ParentId = _ParentId;
            this.GroupId = _GroupId;
            this.value = _value;
        }
    }
}
