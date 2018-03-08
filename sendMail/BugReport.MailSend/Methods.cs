using System;
using System.Net;
using System.Net.Mail;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using System.Web;
using Google.Apis.Drive.v3;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using System.Threading;
using Google.Apis.Util.Store;


namespace BugReport.MailSend
{
    struct MailInfo
    {
        public string[] To;
        public string[] Subject;
        public string[] Body;
        public string[] Attachment;
    };
    public class Mails
    {
        int rows;
        MailInfo email;
        //string[] To, Subject, Body, Attachment = new string[rows];
        public void Sent(string subject,string body, string to, string cc)
        {
            var emailAcount = "etrmailserver@gmail.com";//ConfigurationManager.AppSettings["etrtestsmtp@163.com"];
            var emailPassword = "Welcome2Pwc";// ConfigurationManager.AppSettings["teda00"];

            MailMessage message = new MailMessage();
            //设置发件人,发件人需要与设置的邮件发送服务器的邮箱一致
            MailAddress fromAddr = new MailAddress("etrmailserver@gmail.com");
            message.From = fromAddr;
            //设置收件人,可添加多个,添加方法与下面的一样
            message.To.Add(to);
            //设置抄送人
            message.CC.Add(cc);
            //设置邮件标题
            message.Subject = subject ;
            //设置邮件内容
            message.Body = body;
            //设置邮件发送服务器,服务器根据你使用的邮箱而不同,可以到相应的 邮箱管理后台查看,下面是QQ的
            SmtpClient client = new SmtpClient("smtp.gmail.com", 25);
            //设置发送人的邮箱账号和密码
            client.Credentials = new NetworkCredential(emailAcount, emailPassword);
            //启用ssl,也就是安全发送
            client.EnableSsl = true;
            //发送邮件
            client.Send(message);
        }
        public void  ExcelOpt(string strFileName)
        {
            //string strFileName = "C:/Users/czhang447/Documents/test_VBAmail.xlsx";
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application 
            if (excel == null)
            {
                //this.label1.Text = "Can't access excel";
            }
            else
            {
                excel.Visible = false; excel.UserControl = true;
                // 以只读的形式打开EXCEL文件 
                Excel.Workbook wb = excel.Application.Workbooks.Open(strFileName, missing, true, missing, missing, missing,
                 missing, missing, missing, true, missing, missing, missing, missing, missing);
                //取得第一个工作薄 
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                //取得总记录行数    (包括标题列) 
                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数 
                                                             //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//得到列数 
                                                             //取得数据范围区域   (不包括标题列)   
                this.rows = rowsint;
                Excel.Range rng1 = ws.Cells.get_Range("A2", "A" + rowsint);
                Excel.Range rng2 = ws.Cells.get_Range("B2", "B" + rowsint);
                Excel.Range rng3 = ws.Cells.get_Range("C2", "C" + rowsint);
                Excel.Range rng4 = ws.Cells.get_Range("D2", "D" + rowsint);
                /////receivers' addresses
                string[] toList = new string[rowsint - 1];
                int i = 0;
                foreach (string s in rng1.Value2)
                {
                    Console.WriteLine("sendTo {0}", s);
                    
                    toList[i] = s;//将新值赋给一个数组 
                    i++;
                }
                //Mail's Subject
                string[] subList = new string[rowsint - 1];
                i = 0;
                foreach (string sub in rng2.Value2)
                {
                    subList[i] = sub;
                    i++;
                }
                //content 
                string[] contentList=new string[rowsint - 1];
                i = 0;
                foreach (string con in rng3.Value2)
                {
                    contentList[i] = con;
                    i++;
                }
                //attachment
                string[] attList = new string[rowsint - 1];
                i = 0;
                foreach (string att in rng4.Value2)
                {
                    attList[i] = att;
                    i++;
                }
                
                email.To = toList;
                email.Subject = subList;
                email.Body = contentList;
                email.Attachment = attList;

            }
            excel.Quit(); excel = null;
            Process[] procs = Process.GetProcessesByName("excel");
            foreach (Process pro in procs)
            {
                pro.Kill();//没有更好的方法,只有杀掉进程 
            }
            GC.Collect();
            
        }
        public string[] getTo()
        {
            return email.To;
        }
        public string[] getSubject()
        {
            return email.Subject;
        }
        public string[] getBody()
        {
            return email.Body;
        }
        public string[] getAttachment()
        {
            return email.Attachment;
        }
        public int ExcelRows(string strFileName)
        {
            int rowsint = 0;
            //string strFileName = "C:/Users/czhang447/Documents/test_VBAmail.xlsx";
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application 
            if (excel == null)
            {
                //this.label1.Text = "Can't access excel";
            }
            else
            {
                excel.Visible = false; excel.UserControl = true;
                // 以只读的形式打开EXCEL文件 
                Excel.Workbook wb = excel.Application.Workbooks.Open(strFileName, missing, true, missing, missing, missing,
                 missing, missing, missing, true, missing, missing, missing, missing, missing);
                //取得第一个工作薄 
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                //取得总记录行数    (包括标题列) 
                rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数 
             
            }
            excel.Quit(); excel = null;
            Process[] procs = Process.GetProcessesByName("excel");
            foreach (Process pro in procs)
            {
                pro.Kill();//没有更好的方法,只有杀掉进程 
            }
            GC.Collect();
            return rowsint;
        }


        public static DriveService AuthenticateOauth(string clientSecretJson, string userName)
        {
            try
            {
                if (string.IsNullOrEmpty(userName))
                    throw new ArgumentNullException("userName");
                if (string.IsNullOrEmpty(clientSecretJson))
                    throw new ArgumentNullException("clientSecretJson");
                if (!File.Exists(clientSecretJson))
                    throw new Exception("clientSecretJson file does not exist.");

                // These are the scopes of permissions you need. It is best to request only what you need and not all of them
                string[] scopes = new string[] { DriveService.Scope.DriveReadonly };         	//View the files in your Google Drive                                                 
                UserCredential credential;
                using (var stream = new FileStream(clientSecretJson, FileMode.Open, FileAccess.Read))
                {
                    string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name);

                    // Requesting Authentication or loading previously stored authentication for userName
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets,
                                                                             scopes,
                                                                             userName,
                                                                             CancellationToken.None,
                                                                             new FileDataStore(credPath, true)).Result;
                }

                // Create Drive API service.
                return new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Drive Oauth2 Authentication Sample"
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Create Oauth2 account DriveService failed" + ex.Message);
                throw new Exception("CreateServiceAccountDriveFailed", ex);
            }
        }

        public static void DownloadFile(Google.Apis.Drive.v3.DriveService service, Google.Apis.Drive.v3.Data.File file, string saveTo)
        {

            var request = service.Files.Export(file.Id, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            var stream = new System.IO.MemoryStream();

            
            // Add a handler which will be notified on progress changes.
            // It will notify on each chunk download and when the
            // download is completed or failed.
            request.MediaDownloader.ProgressChanged += (Google.Apis.Download.IDownloadProgress progress) =>
            {
                switch (progress.Status)
                {
                    case Google.Apis.Download.DownloadStatus.Downloading:
                        {
                            Console.WriteLine(progress.BytesDownloaded);
                            break;
                        }
                    case Google.Apis.Download.DownloadStatus.Completed:
                        {
                            Console.WriteLine("Download complete.");
                            SaveStream(stream, saveTo);
                            break;
                        }
                    case Google.Apis.Download.DownloadStatus.Failed:
                        {
                            Console.WriteLine("Download failed.");
                            break;
                        }
                }
            };
            request.Download(stream);

        }

        private static void SaveStream(System.IO.MemoryStream stream, string saveTo)
        {
            using (System.IO.FileStream file = new System.IO.FileStream(saveTo, System.IO.FileMode.Create, System.IO.FileAccess.Write))
            {
                stream.WriteTo(file);
            }
        }

        

    }
    struct TableHead
    {
        public string[] Title;
        public string[] Description;
        public string[] Steps;
        public string[] TestData;
        public string[] ExpectedRes;
        public string[] ActualRes;
        public string[] TestType;
        public string[] CurrentStat;
    };
    public class ExcelFuction
    {
        TableHead TH;
        //return every failed test unit
        public void getFailed()
        {
            string index_failed="";
            int j = 0;
            for(int i = 0; i < TH.CurrentStat.Length; i++)
            {
                if (TH.CurrentStat[i] == "Failed")
                {
                    index_failed += i.ToString()+",";  //存入failed的case的字符串类型索引，之后把用split分开各个数字并转换为整型
                }
            }
        }
        //access to 'strFileName' excel file ,save the datas into Struct TableHead
        public void ExcelOfTest(string strFileName)
        {
            
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application 
            if (excel == null)
            {
                //this.label1.Text = "Can't access excel";
            }
            else
            {
                excel.Visible = false; excel.UserControl = true;
                // 以只读的形式打开EXCEL文件 
                Excel.Workbook wb = excel.Application.Workbooks.Open(strFileName, missing, true, missing, missing, missing,
                 missing, missing, missing, true, missing, missing, missing, missing, missing);
                //取得第一个工作薄 
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                //取得总记录行数    (包括标题列) 
                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数 
                                                             //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//得到列数 
                                                             //取得数据范围区域   (不包括标题列) 
                Excel.Range rng1 = ws.Cells.get_Range("A2", "A" + rowsint);
                Excel.Range rng2 = ws.Cells.get_Range("B2", "B" + rowsint);
                Excel.Range rng3 = ws.Cells.get_Range("C2", "C" + rowsint);
                Excel.Range rng4 = ws.Cells.get_Range("D2", "D" + rowsint);
                Excel.Range rng5 = ws.Cells.get_Range("E2", "E" + rowsint);
                Excel.Range rng6 = ws.Cells.get_Range("F2", "F" + rowsint);
                Excel.Range rng7 = ws.Cells.get_Range("G2", "G" + rowsint);
                Excel.Range rng8 = ws.Cells.get_Range("H2", "H" + rowsint);
                /////receivers' addresses
                TH.Title = new string[rowsint - 1];
                int i = 0;
                foreach (string s in rng1.Value2)
                {
                    //Console.WriteLine("sendTo {0}", s);
                    TH.Title[i] = s;//将新值赋给一个数组 
                    i++;
                }
                //Mail's Subject
                TH.Description = new string[rowsint - 1];
                i = 0;
                foreach (string sub in rng2.Value2)
                {
                    TH.Description[i] = sub;
                    i++;
                }
                //content 
                TH.Steps = new string[rowsint - 1];
                i = 0;
                foreach (string con in rng3.Value2)
                {
                    TH.Steps[i] = con;
                    i++;
                }
                //attachment
                TH.TestData = new string[rowsint - 1];
                i = 0;
                foreach (string att in rng4.Value2)
                {
                    TH.TestData[i] = att;
                    i++;
                }
                TH.ExpectedRes = new string[rowsint - 1];
                i = 0;
                foreach (string att in rng5.Value2)
                {
                    TH.ExpectedRes[i] = att;
                    i++;
                }
                TH.ActualRes = new string[rowsint - 1];
                i = 0;
                foreach (string att in rng6.Value2)
                {
                    TH.ActualRes[i] = att;
                    i++;
                }
                TH.TestType = new string[rowsint - 1];
                i = 0;
                foreach (string att in rng7.Value2)
                {
                    TH.TestType[i] = att;
                    i++;
                }
                TH.CurrentStat = new string[rowsint - 1];
                i = 0;
                foreach (string att in rng8.Value2)
                {
                    TH.CurrentStat[i] = att;
                    i++;
                }
            }
            excel.Quit(); excel = null;
            Process[] procs = Process.GetProcessesByName("excel");
            foreach (Process pro in procs)
            {
                pro.Kill();//没有更好的方法,只有杀掉进程 
            }
            GC.Collect();

        }
        public string[] getTitle()
        {
            return TH.Title;
        }
        public string[] getDescription()
        {
            return TH.Description;
        }
        public string[] getSteps()
        {
            return TH.Steps;
        }
        public string[] getTestData()
        {
            return TH.TestData;
        }
        public string[] getExpectedRes()
        {
            return TH.ExpectedRes;
        }
        public string[] getActualRes()
        {
            return TH.ActualRes;
        }
        public string[] getTestType()
        {
            return TH.TestType;
        }
        public string[] getCurrentStat()
        {
            return TH.CurrentStat;
        }
    }
}
