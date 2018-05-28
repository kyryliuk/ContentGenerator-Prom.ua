using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Content
{
    [Serializable]
 public   class Catalog
    {
        public string Name { get; set; }
        public List<string> Info { get; set; }
        public Catalog() { }
        public Catalog(string _name, List<string> _info)
        {
            this.Name = _name;
            this.Info = _info;
        }
        public Catalog(string _name)
        {
            this.Name = _name;
        }

        public static void SerialToXml()
        {
               XmlSerializer formatter = new XmlSerializer(typeof(Catalog));
                using (FileStream fs = new FileStream("CatalogContent.xml", FileMode.OpenOrCreate))
                {
                    formatter.Serialize(fs, "&&??????");
                    System.Windows.MessageBox.Show("Объект сериализован");
                }
        }
        public static void Desereal(Catalog co)
        {
            XmlSerializer formatter = new XmlSerializer(typeof(Catalog));
            using (FileStream fs = new FileStream("CatalogContent.xml", FileMode.OpenOrCreate))
            {
                co = (Catalog)formatter.Deserialize(fs);
            }
        }
    }
}
