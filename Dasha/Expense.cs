using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dasha
{
    class Expense
    {
        public Expense(string name, string dim, string price, string warehouse, string type, string description, string k)
        {
            this.Name = name;
            this.Dim = dim;
            this.Price = double.Parse(price);
            this.Warehouse = warehouse;
            this.Type = type;
            this.Description = description;
            this.k = double.Parse(k);
        }

        public string Name { get; set; }
        public string Dim { get; set; }
        public string Description { get; set; }
        public double Price { get; set; }
        public string Warehouse { get; set; }
        public string Type { get; set; }

        public double k { get; set; }
    }
}
