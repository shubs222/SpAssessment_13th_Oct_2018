using System;
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
        ClientContext clientcntx;
        Web webpage;
        List<string> headers = new List<string>();
        List<string> items = new List<string>();
        DataTable tbl;
        string Reason;
        bool Status = false;
        string FileSize;
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;

        public void GetSiteData(string Url, string UserName, SecureString passwrd)
        {
            using (clientcntx = new ClientContext(Url))
            {
                clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);
                webpage = clientcntx.Web;
                clientcntx.Load(webpage);
                try
                {
                    clientcntx.ExecuteQuery();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error while fetching the site details: " + e);

                }
                Console.WriteLine("Share Point Site \n Title: " + webpage.Title + "; URL: " + webpage.Url + "; Description: " + webpage.Description);
                Console.ReadKey();

            }

        }
        //public File GetDocument(string Url, string UserName, SecureString passwrd)
        //{

        //    using (clientcntx = new ClientContext(Url))
        //    {
        //        clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);
        //        List documentlist = clientcntx.Web.Lists.GetByTitle("Documents");
        //        clientcntx.Load(documentlist);
        //        clientcntx.ExecuteQuery();
        //        Console.WriteLine("List title: " + documentlist.Title);
        //        Console.ReadKey();
        //        FieldCollection fc = documentlist.Fields;
        //        clientcntx.Load(fc);
        //        clientcntx.ExecuteQuery();
        //        string fileurl = Url + "/_layouts/15/Doc.aspx?sourcedoc=%7Bd9f22086-cf2d-481a-8d1e-b03fd52ceda7%7D&action=default&uid=%7BD9F22086-CF2D-481A-8D1E-B03FD52CEDA7%7D&ListItemId=5&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod";
        //        File myexcelfile = clientcntx.Web.GetFileByUrl(fileurl);
        //        clientcntx.Load(myexcelfile);
        //        try
        //        {
        //            clientcntx.ExecuteQuery();
        //            Console.WriteLine("Success: " + myexcelfile.Name);

        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine("Error while getting document library: " + ex);
        //        }

        //        return myexcelfile;
        //    }

        //}



        /****************************Method for Getting document library data and upadting excel sheet************************/
        public void GetDocumentData(string Url, string UserName, SecureString passwrd)
        {
            using (clientcntx = new ClientContext(Url))
            {
                clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);
                List list = clientcntx.Web.Lists.GetByTitle("MyDocuments");
                string fileurl = Url + "/_layouts/15/Doc.aspx?sourcedoc=%7Bd9f22086-cf2d-481a-8d1e-b03fd52ceda7%7D&action=default&uid=%7BD9F22086-CF2D-481A-8D1E-B03FD52CEDA7%7D&ListItemId=5&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod";

                File file = clientcntx.Web.GetFileByUrl(fileurl);
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();

                //Spreadsheet myfile = new Spreadsheet();
                //Worksheet Sheet = myfile.Workbook.Worksheets.Add(@"D:\FilePathExcelFile.xlsx");
                //MyApp = new Excel.Application();
                //MyApp.Visible = false;
                //MyBook = MyApp.Workbooks.Open(@"D:\FilePathExcelFile");
                //MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
                //lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                //SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                //var workbook = ExcelFile.Load((@"D:\FilePathExcelFile.xlsx"));
                //var worksheet = workbook.Worksheets.ActiveWorksheet;
                OpenExcelFile();
                clientcntx.Load(file);
                clientcntx.ExecuteQuery();
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                    {
                        if (data != null)
                        {
                            data.Value.CopyTo(mStream);
                            pck.Load(mStream);
                            var ws = pck.Workbook.Worksheets.First();
                            tbl = new DataTable();
                            bool hasHeader = true; // adjust it accordingly( i've mentioned that this is a simple approach)
                            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                            {
                                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                                headers.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                            }
                            var startRow = hasHeader ? 2 : 1;
                            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                            {
                                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                                var row = tbl.NewRow();
                                int count = rowNum;
                                Status = true;
                                foreach (var cell in wsRow)
                                {
                                    if (null != cell.Hyperlink)
                                    {
                                        row[cell.Start.Column - 1] = cell.Hyperlink;
                                        items.Add(cell.Hyperlink.ToString());
                                        if (Status == true)
                                        {
                                            Status = UpdateLibrabryData(cell.Hyperlink.ToString(), cell.Address);
                                        }
                                    }
                                    else
                                    {
                                        row[cell.Start.Column - 1] = cell.Text;
                                        items.Add(cell.Text);
                                        if (Status == true)
                                        {
                                            Status = UpdateLibrabryData(cell.Text, cell.Address);
                                        }
                                    }
                                }

                                //MySheet.Cells[rowNum, 4] = FileSize;
                                if (Status == true)
                                {
                                    //MySheet.Cells[rowNum, 5] = "Success";
                                    //MySheet.Cells[rowNum, 6] = Reason;
                                    
                                    UpdateExcelFile(rowNum,Reason,FileSize,"Success");

                                }
                                else
                                {
                                    //    //Sheet.Rows[cell] = "failed";
                                    //MySheet.Cells[rowNum, 5] = "Failed";
                                    //MySheet.Cells[rowNum, 6] = Reason;
                                    UpdateExcelFile(rowNum, Reason, FileSize, "Failed");
                                }


                                tbl.Rows.Add(row);
                            }

                            Console.WriteLine('1');
                        }
                    }
                }
                //MyBook.Save();
                //MyBook.Close();
                CloseExcelFile();
                UploadFileAgain(Url, UserName, passwrd);
                Console.WriteLine("Done");
            }
        }

        /**********************************open local excel file and update the data**************************************/
        public void OpenExcelFile()
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(@"D:\FilePathExcelFile");
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];
        }
        public void CloseExcelFile()
        {
            MyBook.Save();
            MyBook.Close();
        }

        private void UpdateExcelFile(int rowNum,string reason,string fileSize,string uploadStatus)
        {
            MySheet.Cells[rowNum, 4] = fileSize;
            MySheet.Cells[rowNum, 5] = uploadStatus;
            MySheet.Cells[rowNum, 6] = Reason;
        }


        FileCreationInformation Fcreateinfo;
        File FileToUpload;
        System.IO.FileInfo Fileinfo;


        /****************************************************Update Document library items***************************************/
        private bool UpdateLibrabryData(string text, string  Column)
        {
            try
            {
                List list = clientcntx.Web.Lists.GetByTitle("MyDocuments");

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

                            //string[] splitwhole = text.Split(Convert.ToChar(92));
                            //string last = splitwhole[splitwhole.Length - 1];
                            Fcreateinfo.Url = Fileinfo.Name;
                            Fcreateinfo.Content = System.IO.File.ReadAllBytes(text);
                            Fcreateinfo.Overwrite = true;
                            FileToUpload = list.RootFolder.Files.Add(Fcreateinfo);
                            clientcntx.Load(list);
                            clientcntx.ExecuteQuery();
                            Status = true;
                            ListItem li = FileToUpload.ListItemAllFields;
                            li["File_Type"] = System.IO.Path.GetExtension(text);
                            li.Update();
                            clientcntx.ExecuteQuery();
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
                    FieldChoice choice = clientcntx.CastTo<FieldChoice>(field);
                    clientcntx.Load(choice);
                    clientcntx.ExecuteQuery();
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
                    ListItem Li = FileToUpload.ListItemAllFields;
                    Li["FIle_Status"] = StatusUpload;
                    Li.Update();
                    clientcntx.ExecuteQuery();
                    Status = true;
                }
                if (Column.StartsWith("C") && Status)
                {
                    ListItem li = FileToUpload.ListItemAllFields;
                    li["FileCreatedBy"] = text;
                    li.Update();
                    clientcntx.ExecuteQuery();
                    Status = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while updating the Data: " + ex);
            }
            return Status;
        }

        /***********************************************Download file from sharepoint site****************************************/
        public void DownloadFile(string Url, string UserName, SecureString passwrd)
        {
            try
            {
                using (clientcntx = new ClientContext(Url))
                {
                    clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);
                    List documentlist = clientcntx.Web.Lists.GetByTitle("Documents");
                    //string urlforworksheet = Url + documentlist.GetItemById(5);
                    var listItem = documentlist.GetItemById(5);
                    clientcntx.Load(documentlist);
                    clientcntx.Load(listItem, i => i.File);
                    clientcntx.ExecuteQuery();

                    var fileRef = listItem.File.ServerRelativeUrl;
                    var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientcntx, fileRef);
                    //var fileName = System.IO.Path.Combine(urlforworksheet, (string)listItem.File.Name);
                    var fileName = System.IO.Path.Combine(@"D:\", (string)listItem.File.Name);
                    using (var fileStream = System.IO.File.Create(fileName))
                    {
                        fileInfo.Stream.CopyTo(fileStream);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while downloading the file: " + ex);
            }
        }

        /*******************************************upload file again after making changes*****************************************/

        public void UploadFileAgain(string url, string Username, SecureString password)
        {
            using (clientcntx = new ClientContext(url))
            {
                try
                {
                    clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                    List list = clientcntx.Web.Lists.GetByTitle("Documents");
                    FileCreationInformation fcinfo = new FileCreationInformation();
                    fcinfo.Url = "FilePathExcelFile.xlsx";
                    fcinfo.Content = System.IO.File.ReadAllBytes(@"D:\FilePathExcelFile.xlsx");
                    fcinfo.Overwrite = true;
                    File fileToUpload = list.RootFolder.Files.Add(fcinfo);
                    clientcntx.Load(list);
                    clientcntx.ExecuteQuery();
                    Console.WriteLine("Name is : " + fcinfo.Content);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error while uploading file: " + ex);
                }
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

