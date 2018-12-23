using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Content
{    
    [Serializable]
    public  class ReStruct
    {
        public List<string> Product_Del { get; set; }
        public List<string> Group_Del { get; set; }
        public List<twolist> Group_Edit { get; set; }
        public List<twolist> Product_Redirect { get; set; }
        public List<twolist> Group_Redirect { get; set; }
        public List<threelist> Group_New { get; set; }

        public ReStruct()
        {
        }

        public ReStruct(List<string> product_Del, List<string> group_Del, List<twolist> group_Edit, List<twolist> product_Redirect, List<twolist> group_Redirect, List<threelist> group_New)
        {
            Product_Del = product_Del;
            Group_Del = group_Del;
            Group_Edit = group_Edit;
            Product_Redirect = product_Redirect;
            Group_Redirect = group_Redirect;
            Group_New = group_New;
        }
    }
}
