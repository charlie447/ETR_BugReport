using System;
using System.Net;
using System.Net.Mail;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;
using BugReport.MailSend;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Download;

namespace sendMail
{
    public class MainClass
    {
        static string[] Scopes = { DriveService.Scope.DriveReadonly };
        static string ApplicationName = "Drive API .NET Quickstart";
        //static bool mailSent = false;
        //private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        //{
        //    // Get the unique identifier for this asynchronous operation.
        //    String token = (string)e.UserState;

        //    if (e.Cancelled)
        //    {
        //        Console.WriteLine("[{0}] Send canceled.", token);
        //    }
        //    if (e.Error != null)
        //    {
        //        Console.WriteLine("[{0}] {1}", token, e.Error.ToString());
        //    }
        //    else
        //    {
        //        Console.WriteLine("Message sent.");
        //    }
        //    mailSent = true;
        //}
        public static void Main(string[] args)
        {

            //string strFileName = "C:/Users/czhang447/Documents/test_VBAmail.xlsx";

            //Mails mail = new Mails();
            ////mail.Sent("test again", "i love c#", "861789972@qq.com", "charlie.j.zhang@pwc.com");
            //int len = mail.ExcelRows(strFileName);
            //string[] To= new string[len-1];
            //string[] Subject = new string[len - 1];
            //string[] Body = new string[len - 1];
            //string[] Attachment = new string[len - 1];
            //mail.ExcelOpt(strFileName);
            //To = mail.getTo();
            //Subject = mail.getSubject();
            //Body = mail.getBody();
            //Attachment = mail.getAttachment();

            var service = BugReport.MailSend.Mails.AuthenticateOauth("client_secret.json", "user");
            // download each file
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.PageSize = 10;
            listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                .Files;
            Console.WriteLine("Files:");
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    Console.WriteLine("{0} ({1})", file.Name, file.Id);
                    BugReport.MailSend.Mails.DownloadFile(service, file, string.Format(@"C:\sys\google_files\{0}", file.Name+".xlsx"));
                }
            }
            
           }

        
        
    
    }
}
