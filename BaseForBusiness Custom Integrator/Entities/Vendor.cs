using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BaseForBusinessCustomIntegrator
{
    class Vendor
    {
        public int VendorId = 0;
        public string Name = "";
        public string sAddress1 = string.Empty;
        public string sAddress2 = string.Empty;
        public string sAddress3 = string.Empty;
        public string sAddress4 = string.Empty;
        public string sAddress5 = string.Empty;
        public string sContact = string.Empty;
        public string sPhone1 = string.Empty;
        public string sPhone2 = string.Empty;
        public string sFax = string.Empty;
        public string sEmail = string.Empty;
        public string sNote = string.Empty;
        public int iTermsId = 0;
        public int iCompanyId = 0;
        public string sCompanyName = string.Empty;
        public string sFirstName = string.Empty;
        public string sLastName = string.Empty;

        public Vendor(string name)
        {
            Name = name;
        }

        public Vendor()
        {
        }

        public Vendor(string name, int id)
        {
            Name = name;
            VendorId = id;
        }
    }
}
