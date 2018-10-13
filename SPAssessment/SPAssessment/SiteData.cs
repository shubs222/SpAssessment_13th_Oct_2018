using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Data;
using Microsoft.Office.SharePoint;

namespace SPAssessment
{

    class SiteData
    {
        ClientContext clientcntx;   
        Web webpage;
        List<string> headers = new List<string>();
        List<string> items = new List<string>();
        DataTable tbl;
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
                    Console.WriteLine("Error: " + e);
                    throw e;
                }
                Console.WriteLine("Share Point Site \n Title: " + webpage.Title + "; URL: " + webpage.Url + "; Description: " + webpage.Description);
                Console.ReadKey();

                //Console.WriteLine("Do you want to change the name of the site 1. Yes \t 2. press any key to exit");
                //string answer = Console.ReadLine();
                //if (answer.ToUpper() == "YES")
                //{
                //    string Title;
                //    Console.WriteLine("Enter the title");
                //    Title = Console.ReadLine();
                //    webpage.Title = Title;
                //    webpage.Update();
                //    try
                //    {
                //        clientcntx.ExecuteQuery();
                //    }
                //    catch (Exception e)
                //    {
                //        Console.WriteLine("Error " + e);
                //        throw e;
                //    }
                //    Console.WriteLine("New web title is: " + webpage.Title);
                //    Console.ReadKey();
                //}
                //else
                //{

                //}
            }

        }
        public File GetDocument(string Url, string UserName, SecureString passwrd)
        {
            using (clientcntx = new ClientContext(Url))
            {
                    clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);
                    List documentlist = clientcntx.Web.Lists.GetByTitle("Documents");
                    clientcntx.Load(documentlist);
                    clientcntx.ExecuteQuery();
                    Console.WriteLine("List title: " + documentlist.Title);
                     Console.ReadKey();
                    FieldCollection fc = documentlist.Fields;
                    clientcntx.Load(fc);
                    clientcntx.ExecuteQuery();
                    string fileurl = Url+"/_layouts/15/Doc.aspx?sourcedoc=%7Bd9f22086-cf2d-481a-8d1e-b03fd52ceda7%7D&action=default&uid=%7BD9F22086-CF2D-481A-8D1E-B03FD52CEDA7%7D&ListItemId=5&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod";
                    File myexcelfile = clientcntx.Web.GetFileByUrl(fileurl);
                     clientcntx.Load(myexcelfile);
                    clientcntx.ExecuteQuery();
                    Console.WriteLine("Success: "+myexcelfile.Name);
                    return myexcelfile;
            }

        }

        public void GetFilePath(string Url, string UserName, SecureString passwrd)
        {
            //using (clientcntx = new ClientContext(Url))
            //{
            //    clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);
            //    File paths = GetDocument(Url, UserName, passwrd);
            //    FileInformation information = File.OpenBinaryDirect(clientcntx,paths.ServerRelativeUrl);
            //    //using (System.IO.StreamReader sr = new System.IO.StreamReader(information.Stream))
            //    //{
            //    //    // Read the stream to a string, and write the string to the console.
            //    //   String line = sr.ReadToEnd();
            //    //    Console.WriteLine("data: "+line);
            //    //}
                

            //    System.IO.Stream stream = information.Stream;

            //    using (System.IO.StreamReader sr = new System.IO.StreamReader(stream))
            //    {
            //        while (sr.Peek() >= 0)
            //        {
            //            Console.WriteLine(sr.ReadLine());
            //        }
            //    }
            //}

            using(clientcntx = new ClientContext(Url))
            {
                clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);
                string fileurl = Url + "/_layouts/15/Doc.aspx?sourcedoc=%7Bd9f22086-cf2d-481a-8d1e-b03fd52ceda7%7D&action=default&uid=%7BD9F22086-CF2D-481A-8D1E-B03FD52CEDA7%7D&ListItemId=5&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod";
                File file = clientcntx.Web.GetFileByUrl(fileurl);
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                clientcntx.Load(file);
                clientcntx.ExecuteQuery();
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    //using (var stream = File.OpenRead(""))
                    //{
                    //    pck.Load(stream);
                    //}
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
                                foreach (var cell in wsRow)
                                {
                                    if (null != cell.Hyperlink)
                                    {
                                        row[cell.Start.Column - 1] = cell.Hyperlink;
                                        items.Add(cell.Hyperlink.ToString());
                                    }
                                    else
                                    {
                                        row[cell.Start.Column - 1] = cell.Text;
                                        items.Add(cell.Text);
                                    }
                                }
                                tbl.Rows.Add(row);
                                
                            }
                            Console.WriteLine('1');

                        }
                    }



                }

                Console.WriteLine("Done");
                Console.ReadKey();

            }
        }
        public void UploadData(string Url, string UserName, SecureString passwrd)
        {
            using (clientcntx = new ClientContext(Url))
            {
                clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);
                List list = clientcntx.Web.Lists.GetByTitle("Documents");
               
                for (int FPitems = 0; FPitems < items.Count; FPitems+=4)
                {
                    FileCreationInformation fcinfo = new FileCreationInformation();
                    string whole = items[FPitems];
                    string [] splitwhole = whole.Split(Convert.ToChar(92));
                    string last = splitwhole[splitwhole.Length - 1];
                    fcinfo.Url = last;
                    string path = items[FPitems];
                    fcinfo.Content = System.IO.File.ReadAllBytes(path);
                    fcinfo.Overwrite = true;
                    

                    File fileToUpload = list.RootFolder.Files.Add(fcinfo);
                    clientcntx.Load(list);
                    clientcntx.ExecuteQuery();
                    ListItemCreationInformation itemCreationInformation = new ListItemCreationInformation();
                    ListItem li = list.AddItem(itemCreationInformation);
                    li["FileCreatedBy"] = items[FPitems + 2];

                    li.Update();
                    clientcntx.ExecuteQuery();
                    CreatdBy();
                }


            }
        }

        public void CreatdBy()
        {
            for (int authercount = 2; authercount < items.Count; authercount+=4)
            {
                List list = clientcntx.Web.Lists.GetByTitle("Documents");
                //ListItemCreationInformation itemCreationInformation = new ListItemCreationInformation();
                //FileCollection fc = list.AddItem(itemCreationInformation);
                //li["FileCreatedBy"] = items[authercount];

                //li.Update();
                //clientcntx.ExecuteQuery();

                var creationInformation = new ListItemCreationInformation();
                Microsoft.SharePoint.Client.ListItem listItem = list.AddItem(creationInformation);
                listItem.FieldValues["FileCreatedBy"] =  items[authercount];
                listItem.Update();
                clientcntx.ExecuteQuery();

            }
        }

        //public void UploadFile(string url, string Username, SecureString password)
        //{
        //    using (clientcntx = new ClientContext(url))
        //    {
        //        clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
        //        List list = clientcntx.Web.Lists.GetByTitle("MyDocuments");
        //        FileCreationInformation fcinfo = new FileCreationInformation();
        //        fcinfo.Url = "MyDocuments/NewFiles/Products1.txt";
        //        fcinfo.Content = System.IO.File.ReadAllBytes(@"D:\My Tasks\SharePointPractice\A_8th_Oct2018\Products1.txt");
        //        fcinfo.Overwrite = true;
        //        File fileToUpload = list.RootFolder.Files.Add(fcinfo);
        //        clientcntx.Load(list);
        //        clientcntx.ExecuteQuery();
        //        Console.WriteLine("Name is : " + fcinfo.Content);
        //    }
        //}

    }
}
