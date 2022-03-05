using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BaseForBusinessCustomIntegrator
{
    class Term
    {
        public string Name = "";
        public int TermId = -1;

        public Term(string name)
        {
            Name = name;
        }

        public Term()
        {
        }

        public Term(string name, int id)
        {
            Name = name;
            TermId = id;
        }
    }
}
