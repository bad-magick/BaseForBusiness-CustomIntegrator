using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace BaseForBusinessCustomIntegrator
{
    class Items : IEnumerable
    {
        private List<Item> list = new List<Item>();

        IEnumerator IEnumerable.GetEnumerator()
        {
            return list.GetEnumerator();
        }

        public void Add(Item item)
        {
            list.Add(item);
        }

        public Item Item(int index)
        {
            return list[index];
        }

        public int Count()
        {
            return list.Count();
        }

    }
}
