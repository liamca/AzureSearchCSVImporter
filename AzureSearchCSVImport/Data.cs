using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureSearchCSVImport
{
    public class FieldProperties
    {
        public string FieldType { get; set; }
        public string Retreiveable { get; set; }
        public string Filterable { get; set; }
        public string Sortable { get; set; }
        public string Facetable { get; set; }
        public string Searchable { get; set; }
    }
}
