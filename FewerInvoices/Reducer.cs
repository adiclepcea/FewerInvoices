using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FewerInvoices
{
    class Reducer
    {
        public static List<Invoice> Reduce(List<Invoice> Invoices, List<Item> Items)
        {

            List<Invoice> lstFactFinale = new List<Invoice>();

            while (Items.Find((it) => !it.Visited) != null)
            {
                Invoices.Sort();
                Invoices.Reverse();
                Invoice inv = Invoices.ElementAt(0);

                foreach (Item it in inv.GetUnvisitedItems())
                {
                    it.Visited = true;
                }

                lstFactFinale.Add(inv);
            }

            return lstFactFinale;
        }
    }
}
