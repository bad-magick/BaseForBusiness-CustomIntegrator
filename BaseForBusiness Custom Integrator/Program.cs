using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BaseForBusinessCustomIntegrator
{
    class Program
    {
        private Worker worker = null;

        static void Main(string[] args)
        {
            Worker worker = new Worker(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\excelpkg (2).xls");
            //worker.EraseTableData();
            //worker.ProcessCustomersTable();
            //worker.ProcessVendorsTable();
            //worker.ProcessItemsTable();
            worker.ProcessEmployeesTable();
            worker.Release();

            Console.Write("Press any key to close window...");
            Console.ReadKey(false);

        }
    }
}
