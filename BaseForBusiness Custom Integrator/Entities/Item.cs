using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BaseForBusinessCustomIntegrator
{
    class Item
    {
        public string Name = string.Empty;
        public int ItemId = -1;
        public string UPC = string.Empty;
        public string Stock = string.Empty;
        public string Weight = string.Empty;
        public string Size = string.Empty;
        public string Materials = string.Empty;
        public string Color = string.Empty;
        public string Misc = string.Empty;
        public decimal Cost = 0;
        public decimal Price = 0;
        public string Vendor = string.Empty;
        public DateTime Date = DateTime.MinValue;
        public int VendorId = 0;
        public int RevisionId = 0;
        public string RefNum = string.Empty;

        public Item(string name)
        {
            Name = name;
        }

        public Item()
        {
        }

        public Item(string name, int id)
        {
            Name = name;
            ItemId = id;
        }

    }
}
