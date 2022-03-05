using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BaseForBusinessCustomIntegrator
{
    class Company
    {
        public string Name = "";
        public int CompanyId = -1;

        public string BAddress1 = string.Empty;
        public string BAddress2 = string.Empty;
        public string BAddress3 = string.Empty;
        public string BAddress4 = string.Empty;
        public string BAddress5 = string.Empty;

        public string SAddress1 = string.Empty;
        public string SAddress2 = string.Empty;
        public string SAddress3 = string.Empty;
        public string SAddress4 = string.Empty;
        public string SAddress5 = string.Empty;

        public string Phone1 = string.Empty;
        public string Phone2 = string.Empty;
        public string Fax = string.Empty;

        public Company(string name)
        {
            Name = name;
        }

        public Company()
        {
        }

        public Company(string name, int id)
        {
            Name = name;
            CompanyId = id;
        }
    }
}
