using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertToJson
{
    public class SheetData
    {
        public string name { get; set; }
        public string description { get; set; }
        public string external_url { get; set; }
        public string image { get; set; }
        public string attributes { get; set; }
    }
    public class Attribute
    {
        public string trait_type { get; set; }
        public string value { get; set; }
    }

    public class SheetJsonData
    {
        public string name { get; set; }
        public string description { get; set; }
        public string external_url { get; set; }
        public string image { get; set; }
        public List<Attribute> attributes { get; set; }
    }
}
