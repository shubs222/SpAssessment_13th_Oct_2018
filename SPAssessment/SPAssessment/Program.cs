using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using Microsoft.SharePoint.Client;
namespace SPAssessment
{
    class Program
    {
        static string UserName;
        static SecureString passwrd;
        public static void Getdata()
        {
            //Console.WriteLine("Enter user name");
            UserName = "arvind.torvi@acuvate.com";
            Console.WriteLine("Enter password");
            passwrd = GetPassword();

        }
        private static SecureString GetPassword()
        {
            ConsoleKeyInfo info;
            //Get the user's password as a SecureString  
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }

        static void Main(string[] args)
        {
            Getdata();
            string Url = "https://acuvatehyd.sharepoint.com/teams/shubhamtrial";
            SiteData sitedata = new SiteData();
            sitedata.GetSiteData(Url,UserName,passwrd);
            //sitedata.GetDocument(Url, UserName, passwrd);
            sitedata.GetFilePath(Url, UserName, passwrd);
            //sitedata.DownloadFile(Url, UserName, passwrd);
            //sitedata.UploadData(Url, UserName, passwrd);
            Console.ReadKey();
        }
    }
}
