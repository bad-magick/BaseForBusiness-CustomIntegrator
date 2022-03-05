using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        private Items items = new Items();
        private Vendors vendors = new Vendors();

        private IQClient client = null;
        private IQApplication application = null;

        private bool bUseQB = true;

        //private QClient qbClient = new QClient();

        public Worker(string fileSpec)
        {
            Console.WriteLine("Opening Excel workbook...");
            workbook = excelApp.Workbooks.Open(fileSpec);
            if (bUseQB)
            {
                Console.WriteLine("Logging into QuickBase...");
                client = QuickBase.Login("den.boice@baseforbusiness.com", "lovers4life");

                Console.WriteLine("Connecting to application...");
                application = client.Connect("bhkkqamd8", "c7mzr4hbhxuqtiyc3wrezvhb9g");
                //application = client.Connect("bhkeymmdz", "cfzywf8dfiwyrdba5dceg8a7v");
            }
            else
            {
                Console.WriteLine("QuickBase disabled.");
            }
        }

        public void EraseTableData()
        {
            if (bUseQB)
            {
                foreach (KeyValuePair<string, IQTable> table in application.GetTables())
                {
                    switch (table.Value.TableName.ToLower().Trim())
                    {
                        case "customers":
                            Console.WriteLine("Erasing Table Data [" + table.Value.TableName + "]");
                            table.Value.PurgeRecords();
                            break;
                        case "terms":
                            Console.WriteLine("Erasing Table Data [" + table.Value.TableName + "]");
                            table.Value.PurgeRecords();
                            break;
                        case "adresses":
                            Console.WriteLine("Erasing Table Data [" + table.Value.TableName + "]");
                            table.Value.PurgeRecords();
                            break;
                        case "items":
                            Console.WriteLine("Erasing Table Data [" + table.Value.TableName + "]");
                            table.Value.PurgeRecords();
                            break;
                        case "revisions":
                            Console.WriteLine("Erasing Table Data [" + table.Value.TableName + "]");
                            table.Value.PurgeRecords();
                            break;
                        case "vendors":
                            Console.WriteLine("Erasing Table Data [" + table.Value.TableName + "]");
                            table.Value.PurgeRecords();
                            break;
                        case "employees":
                            Console.WriteLine("Erasing Table Data [" + table.Value.TableName + "]");
                            table.Value.PurgeRecords();
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        public void ProcessCustomersTable ()
        {

            Console.WriteLine("Activating sheet [Customers]...");
            Worksheet sheet = workbook.Sheets["Customers"];
            Range range = sheet.get_Range("A2", "AI1802");

            Console.WriteLine("Checking Terms table...");
            //first get all terms
            foreach (Range row in range.Rows)
            {
                string termName = Convert.ToString(row.Cells[1, 26].Value2);
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
                        Term term = new Term(termName); if (bUseQB)
                        {
                            //Terms table
                            IQTable table = application.GetTable("bhkkqamga");
                            IQRecord record = table.NewRecord();
                            record["Name"] = termName;
                            record["Discount %"] = "0";
                            record.AcceptChanges();
                            term.TermId = record.RecordId;
                        }
                        terms.Add(term);
                    }
                }
            }

            //get all companies
            foreach (Range row in range.Rows)
            {
                string companyName = IfIsNull(row.Cells[1, 2].Value2, "");
                if (string.IsNullOrEmpty(companyName))
                {
                    companyName = string.Empty + "";
                }
                else
                {
                    if (companyName.IndexOf(':') != -1)
                    {
                        companyName = companyName.Substring(0, companyName.IndexOf(':'));
                    }
                }


  
                string BAddress1 = Convert.ToString(row.Cells[1, 8].Value2);
                string BAddress2 = Convert.ToString(row.Cells[1, 9].Value2);
                string BAddress3 = Convert.ToString(row.Cells[1, 10].Value2);
                string BAddress4 = Convert.ToString(row.Cells[1, 11].Value2);
                string BAddress5 = Convert.ToString(row.Cells[1, 12].Value2);

                string SAddress1 = Convert.ToString(row.Cells[1, 13].Value2);
                string SAddress2 = Convert.ToString(row.Cells[1, 14].Value2);
                string SAddress3 = Convert.ToString(row.Cells[1, 15].Value2);
                string SAddress4 = Convert.ToString(row.Cells[1, 16].Value2);
                string SAddress5 = Convert.ToString(row.Cells[1, 17].Value2);

                string Phone1 = Convert.ToString(row.Cells[1, 18].Value2);
                string Phone2 = Convert.ToString(row.Cells[1, 19].Value2);
                string Fax = Convert.ToString(row.Cells[1, 20].Value2);

                string ContactEmail = Convert.ToString(row.Cells[1, 21].Value2);
                string ContactName = Convert.ToString(row.Cells[1, 23].Value2);

                string sTerm = Convert.ToString(row.Cells[1, 26].Value2);
                string Taxable = Convert.ToString(row.Cells[1, 29].Value2);

                Company company = new Company(companyName);

                bool skipCompany = false;
                if (companyName.Trim() != string.Empty)
                {
                    foreach (Company comp in companies)
                    {
                        if (comp.Name == companyName)
                        {
                            skipCompany = true;
                        }
                    }
                }
                else
                {
                    skipCompany = true;
                }

                if (!skipCompany)
                {
                    companies.Add(company);
                    Console.WriteLine("Creating Company: \"" + companyName + "\"");
                    IQRecord record = null;
                    if (bUseQB)
                    {
                        //Customers table
                        IQTable table = application.GetTable("bhkkqameq");
                        record = table.NewRecord();
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
                    }

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
                    {
                        if (bUseQB)
                        {
                            record["Related Term"] = termId.ToString();
                        }
                    }
                    else
                    {
                        if (bUseQB)
                        {
                            record["Related Term"] = "0";
                        }
                    }

                    if (bUseQB)
                    {
                        record.AcceptChanges();
                        company.CompanyId = record.RecordId;

                        //Addresses table
                        IQTable addTable = application.GetTable("bhkkqames");
                        IQRecord addRecord = addTable.NewRecord();
                        addRecord["Address 1"] = IfIsNull(SAddress1, "");
                        addRecord["Address 2"] = IfIsNull(SAddress2, "");
                        addRecord["Address 3"] = IfIsNull(SAddress3, "");
                        addRecord["Address 4"] = IfIsNull(SAddress4, "");
                        addRecord["Address 5"] = IfIsNull(SAddress5, "");
                        addRecord["Related Customer"] = company.CompanyId.ToString();
                        addRecord.AcceptChanges();

                    }
                }
            }

        }

        public void ProcessItemsTable()
        {
            Console.WriteLine("Activating sheet [Items]...");
            Worksheet sheet = workbook.Sheets["Items"];
            Range range = sheet.get_Range("A2", "Q1364");

            int iTotCharsL = 0;
            int iTotCharsR = 0;
            int iGT = 0;
            int iLT = 0;
            int iEq = 0;
            int iRecs = 0;

            foreach (Range row in range.Rows)
            {
                string isHidden = IfIsNull(row.Cells[1, 18].Value2, "Y").ToUpper();
                if (IfIsNull(row.Cells[1, 18].Value2, "Y").ToUpper() == "N")
                {
                    iRecs++;
                    string sDesc = IfIsNull(row.Cells[1, 5].Value2, string.Empty);
                    string sPDesc = IfIsNull(row.Cells[1, 6].Value2, string.Empty);
                    iTotCharsL += sDesc.Length;
                    iTotCharsR += sPDesc.Length;
                    string sDescL = string.Format("{0:0000}", sDesc.Length);
                    string sPDescL = string.Format("{0:0000}", sPDesc.Length);
                    string sActualDesc = string.Empty;

                    //Console.Write(sDescL);
                    if (sDesc.Length > sPDesc.Length)
                    {
                        //Console.Write(" > ");
                        iGT++;
                    }
                    else if (sPDesc.Length > sDesc.Length)
                    {
                        //Console.Write(" < ");
                        iLT++;
                    }
                    else
                    {
                        //Console.Write(" = ");
                        iEq++;
                    }
                    //Console.Write(sPDescL);
                    //Console.Write(" - ");

                    string[] aSplit = null;
                    string[] aSplitter = new string[] { "\\" + "n" };
                    if (sDesc.Length > sPDesc.Length)
                    {
                        aSplit = sDesc.Split(aSplitter, StringSplitOptions.None);
                        sActualDesc = sDesc;

                    }
                    else
                    {
                        aSplit = sPDesc.Split(aSplitter, StringSplitOptions.None);
                        sActualDesc = sPDesc;
                    }

                    int iUnknown = 0;

                    string sUPC = string.Empty;
                    string sStock = string.Empty;
                    string sWeight = string.Empty;
                    string sSize = string.Empty;
                    string sMaterials = string.Empty;
                    string sColor = string.Empty;
                    string sMisc = string.Empty;

                    string sName = IfIsNull(row.Cells[1, 1].Value2, "");
                    string sInviteMType = IfIsNull(row.Cells[1, 4].Value2, "");
                    decimal dCost = Convert.ToDecimal(IfIsNull(row.Cells[1, 13].Value2, "0"));

                    string sCogsAccount = IfIsNull(row.Cells[1, 9].Value2, "");
                    decimal dPrice = Convert.ToDecimal(IfIsNull(row.Cells[1, 12].Value2, "0"));
                    bool bIsPassedThrough = false;

                    string sSizeWeight = string.Empty;
                    string sSizeName = string.Empty;
                    string sRefNum = Convert.ToString(row.Cells[1, 2].Value2);


                    if (IfIsNull(row.Cells[1, 17].Value2, "false") == "Y")
                    {
                        bIsPassedThrough = true;
                    }
                    string sPreferredVendor = IfIsNull(row.Cells[1, 16].Value2, "");
                    bool bTaxable = false;
                    if (IfIsNull(row.Cells[1, 14].Value2, "false") == "Y")
                    {
                        bTaxable = true;
                    }

                    DateTime dDate = CreatedEpoch(Convert.ToInt32(row.Cells[1, 3].Value2));

                    foreach (string s in aSplit)
                    {
                        if (s.Length > 0)
                        {
                            if ((s.Length > 4) && (s.Substring(0, 4).ToUpper() == "UPC#"))
                            {
                                sUPC = s.Substring(4, s.Length - 4).Trim();
                                if (sUPC.IndexOf(" ") >= 10)
                                {
                                    sUPC = sUPC.Substring(0, sUPC.IndexOf(" ")).Trim();
                                }
                            }
                            else if ((s.Length > 7) && (s.Substring(0, 7).ToUpper() == "STOCK #"))
                            {
                                sStock = s.Substring(7, s.Length - 7).Trim();
                            }
                            else if ((s.Length > 2) && ((s.Substring(s.Length - 2) == "oz") || (s.Substring(s.Length - 2) == "lb")))
                            {
                                sWeight = s.Trim(); // s.Trim();
                            }
                            else if (s.IndexOf("oz") != -1)
                            {
                                sSizeWeight = s.ToLower().Substring(0, s.ToLower().IndexOf("oz") + 3).Trim();
                                sSizeName = s.Substring(s.ToLower().IndexOf("oz") + 4).Trim();
                            }
                            else if (s.ToLower().IndexOf("lb") != -1)
                            {
                                sSizeWeight = s.ToLower().Substring(0, s.ToLower().IndexOf("lb") + 3).Trim();
                                sSizeName = s.Substring(s.ToLower().IndexOf("lb") + 4).Trim();
                            }
                            else if (s.IndexOf(" x ") != -1)
                            {
                                sSize = s.Trim();
                            }
                            else if (s.IndexOf("PET") != -1)
                            {
                                sMaterials = s.Trim();
                            }
                            else if (s.ToUpper().IndexOf("COLOR") != -1)
                            {
                                sColor = s.Trim();
                            }
                            else
                            {
                                iUnknown++;
                                if (s.Trim().Length > 0)
                                {
                                    sMisc += s.Trim() + "\\" + "n";
                                }
                            }
                        }

                        if (sMisc.Length > 0)
                        {
                            sMisc = sMisc.Substring(0, sMisc.Length - 2);
                        }
                    }

                    //Console.Write(iUnknown.ToString());

                    //Console.WriteLine();

                    Item item = new Item();
                    item.Color = sColor.Trim();
                    item.Cost = dCost;
                    item.Materials = sMaterials.Trim();
                    item.Misc = sMisc.Trim();
                    item.Name = sName.Trim();
                    item.Price = dPrice;
                    item.Size = sSize.Trim();
                    item.Stock = sStock.Trim();
                    item.UPC = sUPC.Trim();
                    item.Stock = sStock.Trim();
                    item.Vendor = sPreferredVendor.Trim();
                    item.Weight = sWeight.Trim();
                    item.Date = CreatedEpoch(Convert.ToInt32(row.Cells[1, 3].Value2));
                    item.RefNum = Convert.ToString(row.Cells[1, 2].Value2);


                    string sdColors = string.Empty;
                    if (sColor.ToLower().IndexOf("colors") != -1)
                    {
                        string[] colors = null;
                        colors = sColor.Split(' ');
                        sdColors = colors[0];
                        int dTryParse = 0;
                        if (!int.TryParse(sdColors, out dTryParse))
                        {
                            sdColors = "0";
                        }

                    }
                    else
                    {
                        sdColors = "1";
                    }

                    Console.Write("Creating Item: " + item.Name);
                    foreach (Vendor vendor in vendors)
                    {
                        if (vendor.Name.ToUpper() == item.Vendor)
                        {
                            item.VendorId = vendor.VendorId;
                            Console.Write(" - Vendor #" + vendor.VendorId.ToString());
                        }
                    }

                    Console.WriteLine();

                    if (bUseQB)
                    {
                        //Items table
                        IQTable addTable = application.GetTable("bhkkqamfp");
                        IQRecord addRecordItem = addTable.NewRecord();
                        addRecordItem["Name"] = item.Name.Trim();
                        addRecordItem["UPC Code"] = item.UPC.Trim();
                        addRecordItem["Package Size"] = sSizeWeight;
                        addRecordItem["Description"] = sSizeName;
                        addRecordItem["Ref Num"] = item.RefNum;

                        addRecordItem.AcceptChanges();
                        item.ItemId = addRecordItem.RecordId;

                        //Revisions table
                        IQRecord addRecordRev = application.GetTable("bhkmbadhy").NewRecord();
                        addRecordRev["Name"] = "Import Revision - " + sName;
                        addRecordRev["Color"] = sdColors.Trim();
                        addRecordRev["Date"] = dDate.ToLongDateString() + " " + dDate.ToLongTimeString();
                        addRecordRev["Matte"] = "0";
                        addRecordRev["Type"] = "Revised Art";
                        addRecordRev["Format"] = "Rollstock";
                        addRecordRev["Bag Pouch Style"] = "Other";
                        addRecordRev["Reclosable Feature"] = "None";
                        addRecordRev["Roll Style Sheet"] = "Other";
                        addRecordRev["Winding"] = "Other";
                        addRecordRev["Core Size"] = "Other";
                        addRecordRev["Extras"] = "None";
                        addRecordRev["Other"] = "";
                        addRecordRev["Misc"] = sActualDesc;

                        addRecordRev["Cost"] = dCost.ToString();
                        addRecordRev["Price"] = dPrice.ToString();
                        addRecordRev["InviteMType"] = sInviteMType.Trim();
                        addRecordRev["CogsAccount"] = sCogsAccount.Trim();
                        addRecordRev["Taxable"] = TrueFalse(bTaxable);
                        addRecordRev["IsPassedThrough"] = TrueFalse(bIsPassedThrough);
                        //addRecordRev["Related Vendor"] = item.VendorId.ToString();
                        addRecordRev["Related Item"] = item.ItemId.ToString();
                        addRecordRev["Ref Num"] = item.RefNum;
                        addRecordRev.AcceptChanges();
                        //addRecordRev["Preferred Vendor"] = sPreferredVendor;
//                        System.Threading.Thread.Sleep(5000);
  //                      addRecordRev.AcceptChanges();
                        item.RevisionId = addRecordRev.RecordId;

                    }

                    items.Add(item);
                }
            }

            //Console.WriteLine();
            //Console.WriteLine("LT: " + iLT.ToString());
            //Console.WriteLine("GT: " + iGT.ToString());
            //Console.WriteLine("Eq: " + iEq.ToString());
            //Console.WriteLine("Avg L: " + (iTotCharsL / iRecs).ToString());
            //Console.WriteLine("Avg R: " + (iTotCharsR / iRecs).ToString());

        }

        public void ProcessVendorsTable()
        {
            Console.WriteLine("Activating sheet [Vendors]...");
            Worksheet sheet = workbook.Sheets["Vendors"];
            Range range = sheet.get_Range("A2", "U1015");
            IQTable addTable = null;

            if (bUseQB)
            {
                //Vendors table
                Console.WriteLine("Opening [Vendors] table...");
                addTable = application.GetTable("bhkkqame8");
            }

            foreach (Range row in range.Rows)
            {
                Vendor vendor = new Vendor();
                vendor.sAddress1 = IfIsNull(row.Cells[1, 5].Value2, "");
                vendor.sAddress2 = IfIsNull(row.Cells[1, 6].Value2, "");
                vendor.sAddress3 = IfIsNull(row.Cells[1, 7].Value2, "");
                vendor.sAddress4 = IfIsNull(row.Cells[1, 8].Value2, "");
                vendor.sAddress5 = IfIsNull(row.Cells[1, 9].Value2, "");
                vendor.Name = IfIsNull(row.Cells[1, 1].Value2, "");
                vendor.sCompanyName = IfIsNull(row.Cells[1, 18].Value2, "");
                vendor.sContact = IfIsNull(row.Cells[1, 10].Value2, "");
                vendor.sEmail = IfIsNull(row.Cells[1, 14].Value2, "");
                vendor.sFax = IfIsNull(row.Cells[1, 13].Value2, "");
                vendor.sFirstName = IfIsNull(row.Cells[1, 19].Value2, "");
                vendor.sLastName = IfIsNull(row.Cells[1, 21].Value2, "");
                vendor.sNote = IfIsNull(row.Cells[1, 15].Value2, "");
                vendor.sPhone1 = IfIsNull(row.Cells[1, 11].Value2, "");
                vendor.sPhone2 = IfIsNull(row.Cells[1, 12].Value2, "");
                vendors.Add(vendor);

                Console.WriteLine("Creating Vendor: " + vendor.Name);
                if (bUseQB)
                {
                    IQRecord addRecord = addTable.NewRecord();
                    addRecord["Name"] = vendor.Name.Trim();
                    addRecord["Fax"] = vendor.sFax.Trim();
                    addRecord["Phone 1"] = vendor.sPhone1.Trim();
                    addRecord["Address 1"] = vendor.sAddress1.Trim();
                    addRecord["Address 2"] = vendor.sAddress2.Trim();
                    addRecord["Contact"] = vendor.sContact.Trim();
                    addRecord["E-mail"] = vendor.sEmail.Trim();

                    addRecord["Address 3"] = vendor.sAddress3.Trim();
                    addRecord["Address 4"] = vendor.sAddress4.Trim();
                    addRecord["Address 5"] = vendor.sAddress5.Trim();
                    addRecord["Phone 2"] = vendor.sPhone2.Trim();
                    addRecord["Company Name"] = vendor.sCompanyName.Trim();
                    addRecord["Note"] = vendor.sNote.Trim();

                    addRecord.AcceptChanges();
                    vendor.VendorId = Convert.ToInt32(Convert.ToString(addRecord.RecordId));
                }
            }
        }

        public void ProcessEmployeesTable()
        {
            Console.WriteLine("Activating sheet [Employees]...");
            Worksheet sheet = workbook.Sheets["Employees"];
            Range range = sheet.get_Range("A2", "J12");

            foreach (Range row in range.Rows)
            {
                string sFirstName = IfIsNull(row.Cells[1, 1].Value2, "").Trim();
                string sMiddleName = IfIsNull(row.Cells[1, 2].Value2, "").Trim();
                string sLastName = IfIsNull(row.Cells[1, 3].Value2, "").Trim();
                string sAddress = IfIsNull(row.Cells[1, 5].Value2, "").Trim();
                string sCity = IfIsNull(row.Cells[1, 6].Value2, "").Trim();
                string sState = IfIsNull(row.Cells[1, 7].Value2, "").Trim();
                string sZip = IfIsNull(row.Cells[1, 8].Value2, "").Trim();
                string sPhone1 = IfIsNull(row.Cells[1, 9].Value2, "").Trim();
                string sPhone2 = IfIsNull(row.Cells[1, 10].Value2, "").Trim();
                string sRefNum = IfIsNull(row.Cells[1, 4].Value2, "").Trim();

                Console.WriteLine("Creating Employee: " + sFirstName + " " + sLastName);

                if (bUseQB)
                {
                    IQRecord newRecord = application.GetTable("bhkkqame7").NewRecord();
                    newRecord["Title"] = "";
                    newRecord["First Name"] = sFirstName;
                    newRecord["Middle Initial"] = sMiddleName;
                    newRecord["Last Name"] = sLastName;
                    newRecord["Phone 1"] = sPhone1;
                    newRecord["Phone 2"] = sPhone2;
                    newRecord["Address"] = sAddress;
                    newRecord["City"] = sCity;
                    newRecord["State"] = sState;
                    newRecord["Zip"] = sZip;
                    newRecord["Email"] = "";
                    newRecord["Ref Num"] = sRefNum;

                    newRecord.AcceptChanges();
                }
            }
        }

        public void Release()
        {
            Console.WriteLine("Closing Excel workbook...");
            workbook.Close();

            Console.WriteLine("Closing QuickBase application...");
            if (bUseQB)
            {
                client.Logout();
            }

            Console.WriteLine("Done.\n");
        }

        static string IfIsNull(object varToCheck, string defaultValue)
        {
            if (varToCheck == null)
                return defaultValue;
            else
                return varToCheck.ToString();
        }

        static DateTime CreatedEpoch(int timeStamp)
        {
            return new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(timeStamp);
        }

        static string TrueFalse(bool isTrue)
        {
            if (isTrue)
            {
                return "1";
            }
            else
            {
                return "0";
            }
        }

    }
}
