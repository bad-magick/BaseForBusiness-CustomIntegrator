using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace BaseForBusinessCustomIntegrator
{
    class Terms : IEnumerable
    {
        private List<Term> list = new List<Term>();

        IEnumerator IEnumerable.GetEnumerator()
        {
            return list.GetEnumerator();
        }

        public void Add(Term item)
        {
            list.Add(item);
        }

        public Term Item(int index)
        {
            return list[index];
        }

        public int Count()
        {
            return list.Count();
        }

    }
}
