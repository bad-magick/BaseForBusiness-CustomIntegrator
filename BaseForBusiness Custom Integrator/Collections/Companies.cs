using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace BaseForBusinessCustomIntegrator
{
    class Companies : IEnumerable
    {
        private List<Company> list = new List<Company>();

        IEnumerator IEnumerable.GetEnumerator()
        {
            return list.GetEnumerator();
        }

        public void Add(Company item)
        {
            list.Add(item);
        }

        public Company Item(int index)
        {
            return list[index];
        }

        public int Count()
        {
            return list.Count();
        }

    }
}
