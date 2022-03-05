using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace BaseForBusinessCustomIntegrator
{
    class Vendors : IEnumerable
    {
        private List<Vendor> list = new List<Vendor>();

        IEnumerator IEnumerable.GetEnumerator()
        {
            return list.GetEnumerator();
        }

        public void Add(Vendor item)
        {
            list.Add(item);
        }

        public Vendor Item(int index)
        {
            return list[index];
        }

        public int Count()
        {
            return list.Count();
        }

    }
}
