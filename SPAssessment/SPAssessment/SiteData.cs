﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Data;
using Microsoft.Office.SharePoint;
using Excel = Microsoft.Office.Interop.Excel;
using Bytescout.Spreadsheet;
using System.Diagnostics;
using OfficeOpenXml.Style;
using GemBox.Spreadsheet;

namespace SPAssessment
{

    class SiteData
    {
        ClientContext ClientCntx;
        Web Webpage;
        DataTable Table;
        string Reason;
        bool Status = false;
        string FileSize;
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        bool CheckUSer;
        UserCollection SiteUsers;
        public bool GetSiteData(string Url)
        {
            try
            {
                using (ClientCntx = new ClientContext(Url))
                {
                    try
                    {
                        ClientCntx.Credentials = new SharePointOnlineCredentials(UserCredentials.UserName, UserCredentials.Passwrd);
                        CheckUSer = true;
                        try
                        {
                            Webpage = ClientCntx.Web;
                            ClientCntx.Load(Webpage);
                            ClientCntx.ExecuteQuery();
                            Console.WriteLine("Share Point Site \n Title: " + Webpage.Title + "; URL: " + Webpage.Url + "; Description: " + Webpage.Description);
                            Console.ReadKey();
                        }
                        catch (Exception Exceptions)
                        {
                            CheckUSer = false;
                            Console.WriteLine("Error while fetching the site details ");
                            WriteToLog.WriteToLogs(Exceptions);
                        }
                    }
                    catch (Exception Exceptions)
                    {
                        CheckUSer = false;
                        Console.WriteLine("Check user name and password");
                        WriteToLog.WriteToLogs(Exceptions);

                    }
                }
            }
            catch (Exception Exceptions)
            {
                CheckUSer = false;
                Console.WriteLine("Site Url not found: ");
                WriteToLog.WriteToLogs(Exceptions);
            }
            return CheckUSer;
        }
       
        /****************************Method for Getting document library data and upadting excel sheet************************/
        public void GetDocumentData(string Url)
        {
            try
            {
               
                    string ListName = "MyDocuments";
                    List list = ClientCntx.Web.Lists.GetByTitle(ListName);
                    string fileurl = Url + "/_layouts/15/Doc.aspx?sourcedoc=%7Bd9f22086-cf2d-481a-8d1e-b03fd52ceda7%7D&action=default&uid=%7BD9F22086-CF2D-481A-8D1E-B03FD52CEDA7%7D&ListItemId=5&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod";

                    File file = ClientCntx.Web.GetFileByUrl(fileurl);
                    ClientResult<System.IO.Stream> data = file.OpenBinaryStream();

                    OpenExcelFile();
                    ClientCntx.Load(file);
                    ClientCntx.ExecuteQuery();
                    using (var pck = new OfficeOpenXml.ExcelPackage())
                    {
                    using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                    {
                        if (data != null)
                        {
                            data.Value.CopyTo(mStream);
                            pck.Load(mStream);
                            var ws = pck.Workbook.Worksheets.First();
                            Table = new DataTable();
                            bool hasHeader = true;
                            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                            {
                                Table.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));

                            }
                            var startRow = hasHeader ? 2 : 1;

                            GetUsers();
                            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                            {
                                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                                var row = Table.NewRow();
                                int count = rowNum;
                                Status = true;
                                foreach (var cell in wsRow)
                                {
                                    //if (null != cell.Hyperlink)
                                    //{
                                    //    row[cell.Start.Column - 1] = cell.Hyperlink;

                                    //    if (Status == true)
                                    //    {
                                    //        Status = UpdateLibrabryData(cell.Hyperlink.ToString(), cell.Address);
                                    //    }
                                    //}
                                    //else
                                    //{
                                        row[cell.Start.Column - 1] = cell.Text;

                                        if (Status == true)
                                        {
                                            Status = UpdateLibrabryData(cell.Text, cell.Address);
                                        }
                                    //}
                                }
                                if (Status == true)
                                {
                                    UpdateExcelFile(rowNum, Reason, FileSize, "Success");
                                }
                                else
                                {
                                    UpdateExcelFile(rowNum, Reason, FileSize, "Failed");
                                }
                                Table.Rows.Add(row);
                            }
                            Console.WriteLine('1');
                        }
                    }
                    

                    CloseExcelFile();
                    UploadFileAgain(Url);
                    Console.WriteLine("All Done");
                }

            }
            catch (Exception Exceptions)
            {
                Console.WriteLine("Error while getting the excel data from Sharepoint site");
                WriteToLog.WriteToLogs(Exceptions);
            }
        }

        /**********************************open local excel file and update the data**************************************/
        public void OpenExcelFile()
        {
            try
            {
                MyApp = new Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(@"D:\FilePathExcelFile");
                MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            }
            catch (Exception Exceptions)
            {
                Console.WriteLine("Error while opening the excel file ");
                WriteToLog.WriteToLogs(Exceptions);
            }
        }


        //**********************************save and close local excel file and update the data**************************************//
        public void CloseExcelFile()
        {
            MyBook.Save();
            MyBook.Close();
        }

        /**********************************update local excel file and update the data**************************************/
        private void UpdateExcelFile(int rowNum,string reason,string fileSize,string uploadStatus)
        {
            MySheet.Cells[rowNum, 4] = fileSize;
            MySheet.Cells[rowNum, 5] = uploadStatus;
            MySheet.Cells[rowNum, 6] = Reason;
        }


        FileCreationInformation Fcreateinfo;
        File FileToUpload;
        System.IO.FileInfo Fileinfo;
        ListItem Listitem;

        /****************************************************Update Document library items***************************************/
        private bool UpdateLibrabryData(string text, string  Column)
        {
            try
            {
                List list = ClientCntx.Web.Lists.GetByTitle("MyDocuments");

                if (Column.StartsWith("A"))
                {
                    Fileinfo = new System.IO.FileInfo(text);

                    if (Fileinfo.Exists)
                    {

                        double Filesize = (Fileinfo.Length / 1e+6);
                        FileSize = Filesize + "mb";
                        Console.WriteLine("size: " + Filesize);
                        if (Fileinfo.Length < 2000000 && Fileinfo.Length > 100000)
                        {
                            
                            Fcreateinfo = new FileCreationInformation();
                            Fcreateinfo.Url = Fileinfo.Name;
                            Fcreateinfo.Content = System.IO.File.ReadAllBytes(text);
                            Fcreateinfo.Overwrite = true;
                            FileToUpload = list.RootFolder.Files.Add(Fcreateinfo);
                            ClientCntx.Load(list);
                            ClientCntx.ExecuteQuery();
                            Status = true;
                            Listitem = FileToUpload.ListItemAllFields;
                            Listitem["File_Type"] = System.IO.Path.GetExtension(text);
                            Listitem.Update();
                            ClientCntx.ExecuteQuery();
                            Reason = "NA";
                            Console.WriteLine("File : {0} uploaded Successfully", Fileinfo.Name);
                        }
                        else
                        {

                            if (Fileinfo.Length < 100000)
                            {
                                Reason = "File size is Less than Required file size";
                                Console.WriteLine(Reason);
                                Status = false;
                            }
                            else
                            {
                                Reason = "File size is more than Required file size";
                                Console.WriteLine(Reason);
                                Status = false;
                            }

                        }

                    }
                    else
                    {
                        Reason = "File Does not exist";
                        Console.WriteLine(Reason);
                        Status = false;
                    }
                }

                if (Column.StartsWith("B") && Status)
                {
                    Field field = list.Fields.GetByTitle("FIle_Status");
                    FieldChoice choice = ClientCntx.CastTo<FieldChoice>(field);
                    ClientCntx.Load(choice);
                    ClientCntx.ExecuteQuery();
                    string[] MyStatus = text.ToUpper().Split(',');
                    string StatusUpload = string.Empty;
                    for (int choicecount = 0; choicecount < MyStatus.Length; choicecount++)
                    {
                        if (choice.Choices.Contains(MyStatus[choicecount].Trim()))
                        {
                            if (choicecount == MyStatus.Length - 1)
                            {
                                StatusUpload = StatusUpload + MyStatus[choicecount];
                            }
                            else
                            {
                                StatusUpload = StatusUpload + MyStatus[choicecount] + ";";
                            }
                        }
                    }
                    Listitem = FileToUpload.ListItemAllFields;
                    Listitem["FIle_Status"] = StatusUpload;
                    Listitem.Update();
                    ClientCntx.ExecuteQuery();
                    Status = true;
                }
                if (Column.StartsWith("C") && Status)
                {
                    try
                    {
                        User user = SiteUsers.GetByEmail(text);
                        ClientCntx.Load(user);
                        ClientCntx.ExecuteQuery();
                        Listitem = FileToUpload.ListItemAllFields;
                        Listitem["FileCreatedBy"] = user.Title;
                        Listitem.Update();
                        ClientCntx.ExecuteQuery();
                        Status = true;
                    }
                    catch (Exception userexe)
                    {
                        Reason = "User not found";
                        Console.WriteLine();
                        WriteToLog.WriteToLogs(userexe);
                        Status = false;
                        FileToUpload.DeleteObject();
                    }

                }
                
            }
            catch (Exception Exceptions)
            {
                Console.WriteLine("Error while updating the Data: ");
                WriteToLog.WriteToLogs(Exceptions);
            }
            return Status;
        }

        /***********************************************Download file from sharepoint site****************************************/
        public void DownloadFile(string Url)
        {
            try
            {
                    string ListName = "Documents";
                    List documentlist = ClientCntx.Web.Lists.GetByTitle(ListName);
                    string urlforworksheet = Url + documentlist.GetItemById(5);
                    var ListItem = documentlist.GetItemById(5);
                    ClientCntx.Load(documentlist);
                    ClientCntx.Load(ListItem, i => i.File);
                    ClientCntx.ExecuteQuery();
                    var FileRef = ListItem.File.ServerRelativeUrl;
                    var FileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ClientCntx, FileRef);
                    //var fileName = System.IO.Path.Combine(urlforworksheet, (string)listItem.File.Name);
                    var FileName = System.IO.Path.Combine(@"D:\", (string)ListItem.File.Name);
                    using (var fileStream = System.IO.File.Create(FileName))
                    {
                        FileInfo.Stream.CopyTo(fileStream);
                    }
            }
            catch (Exception Exceptions)
            {
                Console.WriteLine("Error while downloading the file: ");
                WriteToLog.WriteToLogs(Exceptions);
            }
        }

        /*******************************************upload file again after making changes*****************************************/

        public void UploadFileAgain(string url)
        {
            try
            {

                List list = ClientCntx.Web.Lists.GetByTitle("Documents");
                FileCreationInformation Fcinfo = new FileCreationInformation();
                Fcinfo.Url = "FilePathExcelFile.xlsx";
                Fcinfo.Content = System.IO.File.ReadAllBytes(@"D:\FilePathExcelFile.xlsx");
                Fcinfo.Overwrite = true;
                File FileToUpload = list.RootFolder.Files.Add(Fcinfo);
                ClientCntx.Load(list);
                ClientCntx.ExecuteQuery();
                Console.WriteLine("Name is : " + Fcinfo.Content);
            }
            catch (Exception Exceptions)
            {
                Console.WriteLine("Error while uploading file: ");
                WriteToLog.WriteToLogs(Exceptions);
            }
        }


        public void GetUsers()
        {
            SiteUsers = ClientCntx.Web.SiteUsers;

            try
            {
                ClientCntx.Load(SiteUsers);
                ClientCntx.ExecuteQuery();
            }
            catch (Exception Exceptions)
            {
                Console.WriteLine("Error while getting site users");
                WriteToLog.WriteToLogs(Exceptions);
            }

        }    
        

    }
}


/****************************************************Previous Code*****************************************************/
//int lastRow;
//public void UpdateExcelData()
//{

//        MyApp = new Excel.Application();
//        MyApp.Visible = false;
//        MyBook = MyApp.Workbooks.Open(@"D:\FilePathExcelFile.xlsx");
//        MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
//        lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
//    MySheet.Cells[lastRow, 1] = emp.Name;

//}







//   Create new Spreadsheet
//Spreadsheet document = new Spreadsheet();

//   // add new worksheet
//   Microsoft.Office.Interop.Excel.Worksheet Sheet = document.Workbook.Worksheets.Add("FormulaDemo");

//            // headers to indicate purpose of the column
//            Sheet.Cell("A1").Value = "Formula (as text)";
//            // set A column width
//            Sheet.Columns[0].Width = 250;

//            Sheet.Cell("B1").Value = "Formula (calculated)";
//            // set B column width
//            Sheet.Columns[1].Width = 250;


//            // write formula as text 
//            Sheet.Cell("A2").Value = "7*3+2";
//            // write formula as formula
//            Sheet.Cell("B2").Value = "=7*3+2";

//            // delete output file if exists already
//            if (File.Exists("Output.xls"))
//            {
//                File.Delete("Output.xls");
//            }

//            // Save document
//            document.SaveAs("Output.xls");

//            // Close Spreadsheet
//            document.Close();

//            // open generated XLS document in default program
//            Process.Start("Output.xls");

//        }

//    }
//}



//public void UploadData(string Url, string UserName, SecureString passwrd)
//{
//    using (clientcntx = new ClientContext(Url))
//    {
//        clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);
//        List list = clientcntx.Web.Lists.GetByTitle("MyDocuments");

//        for (int FPitems = 0; FPitems < items.Count; FPitems+=4)
//        {
//            System.IO.FileInfo fi = new System.IO.FileInfo(items[FPitems]);
//            Console.WriteLine("size: "+fi.Length);
//            if (!(fi.Length > 15000000))
//            {
//                FileCreationInformation fcinfo = new FileCreationInformation();
//                string whole = items[FPitems];
//                string[] splitwhole = whole.Split(Convert.ToChar(92));
//                string last = splitwhole[splitwhole.Length - 1];
//                fcinfo.Url = last;
//                string path = items[FPitems];
//                fcinfo.Content = System.IO.File.ReadAllBytes(path);
//                fcinfo.Overwrite = true;
//                File fileToUpload = list.RootFolder.Files.Add(fcinfo);
//                clientcntx.Load(list);
//                clientcntx.ExecuteQuery();
//ListItemCreationInformation itemCreationInformation = new ListItemCreationInformation();

//                clientcntx.ExecuteQuery();  

//            }
//            //CreatdBy();
//            else
//            {
//                Console.WriteLine("cant insert data");
//            }
//        }

//        //SPWeb web = new SPSite(/*your web url or variable*/).OpenWeb();
//        //SPDocumentLibrary docLib = (SPDocumentLibrary)web.Lists[/*here your document library*/];
//        //docLib.Fields.Add("columName1", SPFieldType.Text, false);
//    }
//}

//public void CreatdBy()
//{
//    for (int authercount = 2; authercount < items.Count; authercount+=4)
//    {
//        List list = clientcntx.Web.Lists.GetByTitle("MyDocuments");
//        //ListItemCreationInformation itemCreationInformation = new ListItemCreationInformation();
//        //FileCollection fc = list.AddItem(itemCreationInformation);
//        //li["FileCreatedBy"] = items[authercount];

//        //li.Update();
//        //clientcntx.ExecuteQuery();

//        var creationInformation = new ListItemCreationInformation();
//        Microsoft.SharePoint.Client.ListItem listItem = list.AddItem(creationInformation);
//        listItem.FieldValues["FileCreatedBy"] = items[authercount];
//        listItem.Update();
//        clientcntx.ExecuteQuery();

//    }
//}

//public void GetLookupValue(string Url, string UserName, SecureString passwrd)
//{
//    using (clientcntx = new ClientContext(Url))
//    {
//        clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);

//        ListItem item = clientcntx.Web.Lists.GetByTitle("Department").GetItemById(1);

//        clientcntx.Load(item);
//        clientcntx.ExecuteQuery();

//        FieldLookupValue lookup = item["Department_Name"] as FieldLookupValue;
//        string lvalue = lookup.LookupValue;
//        int lId = lookup.LookupId;
//    }
//}

//public void SetLookupValue(string Url, string UserName, SecureString passwrd)
//{
//    using (clientcntx = new ClientContext(Url))
//    {
//        clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);

//        ListItem item = clientcntx.Web.Lists.GetByTitle("Department").GetItemById(1);

//        clientcntx.Load(item);
//        clientcntx.ExecuteQuery();

//        FieldLookupValue lookup = item["Department_Name"] as FieldLookupValue;
//        lookup.LookupId = 9;
//        item["Department_Name"] = lookup;
//        item.Update();
//        clientcntx.ExecuteQuery();
//    }

//}

