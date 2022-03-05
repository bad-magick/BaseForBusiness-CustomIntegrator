using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Intuit.QuickBase.Core;
using Intuit.QuickBase.Client;
using Microsoft.Office.Interop.Excel;

namespace BaseForBusinessCustomIntegrator
{
    class Worker
    {
        private Application excelApp = new Application();
        private Workbook workbook = null;
        private Terms terms = new Terms();
        private Companies companies = new Companies();

        //private QClient qbClient = new QClient();

        public Worker(string fileSpec)
        {
            Console.WriteLine("Logging into QuickBase...");
            var client = QuickBase.Login("den.boice@baseforbusiness.com", "lovers4life");

            Console.WriteLine("Connecting to application...");
            var application = client.Connect("bhkkqamd8", "c7mzr4hbhxuqtiyc3wrezvhb9g");

            Console.WriteLine("Opening Excel workbook...");
            workbook = excelApp.Workbooks.Open(fileSpec);

            Console.WriteLine("Activating sheet...");
            Worksheet sheet = workbook.Sheets["Customers"];
            Range range = sheet.get_Range("A2", "AI1802");

            Console.WriteLine("Checking Terms table...");
            //first get all terms
            foreach (Range row in range.Rows)
            {
                string termName = Convert.ToString(row.Cells[1, 22].Value2);
                if (termName == null)
                {
                    termName = string.Empty;
                }
                if (termName.Trim().Length > 0)
                {

                    bool termExists = false;
                    foreach (Term term in terms)
                    {
                        if (term.Name == termName)
                        {
                            termExists = true;
                        }
                    }

                    if (!termExists)
                    {
                        Console.WriteLine("Creating Term: \"" + termName + "\"");
                        Term term = new Term(termName);
                        IQTable table = application.GetTable("bhkkqamga");
                        IQRecord record = table.NewRecord();
                        record["Name"] = termName;
                        record["Discount %"] = "0";
                        record.AcceptChanges();
                        term.TermId = record.RecordId;
                        terms.Add(term);
                    }
                }
            }

            //get all companies
            range = sheet.get_Range("A2", "AI1802");
            foreach (Range row in range.Rows)
            {
                string sName = Convert.ToString(row.Cells[1, 1].Value2);
                string companyName = row.Cells[1, 31].Value2;

                string BAddress1 = Convert.ToString(row.Cells[1, 4].Value2);
                string BAddress2 = Convert.ToString(row.Cells[1, 5].Value2);
                string BAddress3 = Convert.ToString(row.Cells[1, 6].Value2);
                string BAddress4 = Convert.ToString(row.Cells[1, 7].Value2);
                string BAddress5 = Convert.ToString(row.Cells[1, 8].Value2);

                string SAddress1 = Convert.ToString(row.Cells[1, 9].Value2);
                string SAddress2 = Convert.ToString(row.Cells[1, 10].Value2);
                string SAddress3 = Convert.ToString(row.Cells[1, 11].Value2);
                string SAddress4 = Convert.ToString(row.Cells[1, 12].Value2);
                string SAddress5 = Convert.ToString(row.Cells[1, 13].Value2);

                string Phone1 = Convert.ToString(row.Cells[1, 14].Value2);
                string Phone2 = Convert.ToString(row.Cells[1, 15].Value2);
                string Fax = Convert.ToString(row.Cells[1, 16].Value2);

                string ContactEmail = Convert.ToString(row.Cells[1, 17].Value2);
                string ContactName = Convert.ToString(row.Cells[1, 19].Value2);

                string sTerm = Convert.ToString(row.Cells[1, 22].Value2);
                string Taxable = Convert.ToString(row.Cells[1, 25].Value2);

                

                //Console.WriteLine(companyName);
                Company company = new Company(companyName);

                bool foundCompany = false;
                if (!IfIsNull(sName, ":").Contains(':'))
                {
                    if (companyName != "")
                    {
                        if (companyName != null)
                        {
                            foreach (Company comp in companies)
                            {
                                if (comp.Name == companyName)
                                {
                                    foundCompany = true;
                                }
                            }
                        }
                        else
                        {
                            foundCompany = true;
                        }
                    }
                    else
                    {
                        foundCompany = true;
                    }
                }
                else
                {
                    foundCompany = true;
                }

                if (!foundCompany)
                {
                    companies.Add(company);
                    Console.WriteLine("Creating Company: \"" + companyName + "\"");
                    IQTable table = application.GetTable("bhkkqameq");
                    IQRecord record = table.NewRecord();
                    record["Customer Name"] = IfIsNull(companyName, "");
                    record["Contact Name"] = IfIsNull(ContactName, "");
                    record["Contact Email"] = IfIsNull(ContactEmail, "");
                    record["Phone 1"] = IfIsNull(Phone1, "");
                    record["Phone 2"] = IfIsNull(Phone2, "");
                    record["Fax"] = IfIsNull(Fax, "");
                    record["Bill Address 1"] = IfIsNull(BAddress1, "");
                    record["Bill Address 2"] = IfIsNull(BAddress2, "");
                    record["Bill Address 3"] = IfIsNull(BAddress3, "");
                    record["Bill Address 4"] = IfIsNull(BAddress4, "");
                    record["Bill Address 5"] = IfIsNull(BAddress5, "");
                    record["Taxable"] = IfIsNull(Taxable, "0");

                    int termId = -1;
                    if (sTerm == null)
                    {
                        sTerm = string.Empty;
                    }
                    if (sTerm.Trim().Length > 0)
                    {
                        foreach (Term term in terms)
                        {
                            if (sTerm.Trim() == term.Name)
                            {
                                termId = term.TermId;
                            }
                        }
                    }

                    if (termId != -1)
                        record["Related Term"] = termId.ToString();
                    else
                        record["Related Term"] = "0";
                         
                    record.AcceptChanges();
                    company.CompanyId = record.RecordId;

                    IQTable addTable = application.GetTable("bhkkqames");
                    IQRecord addRecord = addTable.NewRecord();
                    addRecord["Address 1"] = IfIsNull(SAddress1, "");
                    addRecord["Address 2"] = IfIsNull(SAddress2, "");
                    addRecord["Address 3"] = IfIsNull(SAddress3, "");
                    addRecord["Address 4"] = IfIsNull(SAddress4, "");
                    addRecord["Address 5"] = IfIsNull(SAddress5, "");
                    addRecord["Related Customer"] = company.CompanyId.ToString();
                    addRecord.AcceptChanges();
                    int RecordId = addRecord.RecordId;

                }
            }

            Console.WriteLine("Closing Excel workbook...");
            workbook.Close();

            Console.WriteLine("Closing QuickBase application...");
            //client.Logout();

            Console.WriteLine("Done.");
            //workbook = excelApp.Workbooks.Open(fileSpec);
            //workbook.Sheets

        }

        static string IfIsNull(string varToCheck, string defaultValue)
        {
            if (varToCheck == null)
                return defaultValue;
            else
                return varToCheck;
        }
    }
}
