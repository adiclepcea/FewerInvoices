using System.Collections.Generic;
using System.Linq;
using System;

namespace FewerInvoices
{

    class Item
    {
        public string Name { get; set; }
        public List<Invoice> Invoices = new List<Invoice>();
        public bool Visited { get; set; }

        public Item(string Name)
        {
            this.Name = Name;
        }

        public void AttemptAdd(Invoice invoice)
        {
            if (!Invoices.Contains(invoice))
            {
                Invoices.Add(invoice);
            }
        }

        public new string ToString()
        {
            return string.Format("{0}: {1}", Name, Invoices.Count);
        }
    }

    class Invoice : IComparable
    {
        public string Name { get; set; }
        public List<Item> Items = new List<Item>();

        public Invoice(string Name)
        {
            this.Name = Name;
        }

        public List<Item> GetUnvisitedItems()
        {
            return Items.Where((i) => !i.Visited).ToList();
        }

        public void AttemptAdd(Item item)
        {
            if (!Items.Contains(item))
            {
                Items.Add(item);
            }
        }

        public int CompareTo(object obj)
        {
            if (obj == null) return 1;
            if (obj.GetType() != this.GetType())
            {
                throw new ArgumentException("Invalid invoice provided");
            }
            return this.GetUnvisitedItems().Count.CompareTo(((Invoice)obj).GetUnvisitedItems().Count);
        }
    }
}
