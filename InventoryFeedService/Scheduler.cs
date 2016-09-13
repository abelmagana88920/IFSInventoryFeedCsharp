using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.IO;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using System.Data.Objects;



namespace InventoryFeedService
{
    public partial class Scheduler : ServiceBase
    {
        private Timer timer1 = null;
        IFSReportingContext db = new IFSReportingContext();

        private static void GetPathParams(out string localPath, out string localPathwofile)
        {
            string path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            //once you have the path you get the directory with:
            var directory = System.IO.Path.GetDirectoryName(path);
            localPath = new Uri(directory).LocalPath;        
            localPathwofile = localPath + "\\..\\..\\App_Data";
        }

        public static void killProcessByName(string processName)
        {
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName(processName);
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }

        public string convertCSVtoXLS(string pathoffile, string localPathwofile, int randomNumber)
        {
            var pathcsv = pathoffile +".csv";
            System.IO.File.Copy(@"" + pathcsv, @"" + pathcsv + ".xls");

            /*Random random = new Random();
            int randomNumber = random.Next(0, 10000); */

            var _app = new Excel.Application();
            var _workbooks = _app.Workbooks;

            _workbooks.OpenText(@"" + pathcsv + ".xls",
                                     DataType: Excel.XlTextParsingType.xlDelimited,
                                     TextQualifier: Excel.XlTextQualifier.xlTextQualifierDoubleQuote,
                                     ConsecutiveDelimiter: true,
                                     Comma: true);
            // Convert To Excle 97 / 2003
            string createdNewFile = pathoffile + ".xls";
            _workbooks[1].SaveAs(createdNewFile, Excel.XlFileFormat.xlExcel5);

            _workbooks.Close();
            _app.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);

            return createdNewFile;
        }


        public string Emails(string customer_no, string filePath, string email_address, IFSReportingContext db_local, tblInventoryFeedProcess inv)
        {          
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                mail.From = new MailAddress("abelmagana88920@gmail.com");
                foreach (var address in email_address.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    Library.WriteErrorLog(address);
                    mail.To.Add(address);
                }
                mail.Subject = "Inventory Feed Request";
                mail.Body = "Inventory Body";

                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(filePath);
                mail.Attachments.Add(attachment);

                SmtpServer.Port = 587;
                SmtpServer.UseDefaultCredentials = false;
                SmtpServer.Credentials = new System.Net.NetworkCredential("abelmagana88920@gmail.com", "magana04426");
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);

                string sendmessage = "Send Email Successfully";
                Library.WriteErrorLog(sendmessage);

                inv.status = "3";
                inv.current_pr = 2;
                inv.datetime_updated = DateTime.Now;
                db_local.SaveChanges();
                
                Library.WriteErrorLog(sendmessage);
                return new JavaScriptSerializer().Serialize(new { message = sendmessage, status = "success" });       
            }
            catch (Exception ex)
            {
                inv.status = "999";
                inv.current_pr = 2;
                db_local.SaveChanges();
                var message_result = ex.Message + ": Failed Message";
                Library.WriteErrorLog(message_result);
                return new JavaScriptSerializer().Serialize(new { message = message_result, status = "failed" });
            }
        }


        public string FTPs(string customer_no, string filePath, string ftp_address, IFSReportingContext db_local, tblInventoryFeedProcess inv)
        {
            string[] separators = { ",", ";", " " };
            string value = ftp_address;
            string[] words = value.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            string ftphostname="", ftpusername="", ftppassword="", ftpfolder="";
           
            if (words.Length <= 4) // ftp host, username, password
            {  
                ftphostname = words[0];
                ftpusername = words[1];
                ftppassword = words[2];
                if (words.Length == 4)
                    ftpfolder = words[3];
                else
                    ftpfolder = "";
            }
          
            Library.WriteErrorLog("FTP HostName: " + ftphostname);
            Library.WriteErrorLog("Username: "+ftpusername);
            Library.WriteErrorLog("Password: " + ftppassword);
            Library.WriteErrorLog("FTPFolder: "+ftpfolder);
            
            string fileName = "";
            Uri uri = new Uri(filePath);
            if (uri.IsFile)
                fileName = System.IO.Path.GetFileName(uri.LocalPath);
            
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://" + ftphostname + "/"+ftpfolder+"/" + fileName);
                request.Method = WebRequestMethods.Ftp.UploadFile;
                request.Credentials = new NetworkCredential(ftpusername, ftppassword);

                // Copy the contents of the file to the request stream.
                // StreamReader sourceStream = new StreamReader("C:\\Users\\Abel\\Desktop\\test.xls");
                //byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
                Library.WriteErrorLog(filePath);
                byte[] fileContents = System.IO.File.ReadAllBytes(filePath);
                // sourceStream.Close();
                request.ContentLength = fileContents.Length;
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(fileContents, 0, fileContents.Length);
                requestStream.Close();

                inv.status = "3";
                inv.current_pr = 2;
                inv.datetime_updated = DateTime.Now;
                db_local.SaveChanges();

                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                response.Close();

                
                var message_result = "Upload File Complete, status";
                Library.WriteErrorLog(message_result);
                return new JavaScriptSerializer().Serialize(new { message = message_result, status = "success" });
               
            }
            catch (Exception ex)
            {
                inv.status = "999";
                inv.current_pr = 2;
                db_local.SaveChanges();
                var message_result = ex.Message + ": FTP";
                Library.WriteErrorLog(message_result);
                return new JavaScriptSerializer().Serialize(new { message = message_result, status = "failed" });
            }
        }


        public string DataCSVXLS(string customer_no, string filetyperequested, string fields ,IFSReportingContext db_local, tblInventoryFeedProcess inv)
        {
            string localPath, localPathwofile;
            string createdNewFile;
            var index=0;
            GetPathParams(out localPath, out localPathwofile);

            try
            {
                var emList = db.tblInvoiceLinesMasters.Select(m => new { m.PART_NO, m.INVOICED_QTY, m.CUSTOMER_NO }).Where(m => m.CUSTOMER_NO == customer_no).ToList();
                if (emList.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    //header
                    string[] SplitString = fields.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                    string[] sArr = new string[SplitString.Length];

                     index=0;
                     foreach (var field in fields.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                     {
                         sArr[index] = string.Format(@"""{0}""", field);
                         index++;
                    }


                     sb.AppendLine(string.Join(",",sArr ));

                     sArr = new string[SplitString.Length]; //resetting again
                    
                     foreach (var i in emList)
                     {
                         index = 0;
                         foreach (var field in fields.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                         {
                             if (field == "PART_NO") sArr[index] = string.Format(@"""{0}""", i.PART_NO);
                             else if (field == "INVOICED_QTY") sArr[index] = string.Format(@"""{0}""", i.INVOICED_QTY);
                             index++;
                         }
                         sb.AppendLine(string.Join(",", sArr));
                     }

 

                    //creating csv file

                    Random random = new Random();
                    int randomNumber = random.Next(0, 10000);
                    var pathoffile = localPathwofile + "\\"+ DateTime.Now.ToString("yyyy'-'MM'-'dd_HH'.'mm'.'ss.fff") + "InventoryFeed" + randomNumber;
                    var pathcsv = pathoffile+ ".csv";

                    Library.WriteErrorLog(pathcsv);
                    using (System.IO.StreamWriter file = new
                    System.IO.StreamWriter(pathcsv))
                    {
                        file.WriteLine(sb.ToString());
                        createdNewFile = pathcsv;
                    }
                    // Convert CSV To xls

                    if (filetyperequested == "XLS")
                    {
                        createdNewFile = convertCSVtoXLS(pathoffile, localPathwofile, randomNumber); //static function

                        killProcessByName("Excel"); //static function
                        string fileName = pathcsv + ".xls";
                        if (fileName != null || fileName != string.Empty) {
                            if ((System.IO.File.Exists(fileName))){
                                System.IO.File.Delete(fileName);
                            }
                        }
                    }

                    inv.status = "2"; 
                    db_local.SaveChanges();
                    var message_result = "Create File Complete";
                    Library.WriteErrorLog(message_result);
                  
                    return new JavaScriptSerializer().Serialize(new { message = message_result, filePath = createdNewFile, status="success" });
                
                }
                else
                {
                    var message_result = "No Fetch";
                    Library.WriteErrorLog(message_result);
                    return new JavaScriptSerializer().Serialize(new { message = message_result, status="0" });
                }
            }
            catch(Exception e)
            {
                inv.status = "999";
                db_local.SaveChanges();
                var message_result = e.Message + ": Not created";
                Library.WriteErrorLog(message_result);
                return new JavaScriptSerializer().Serialize(new { message = message_result, status="failed" });  
            }
        }


        public Scheduler()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            timer1 = new Timer();
            this.timer1.Interval = 5000;
            this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Tick);
            timer1.Enabled = true;
            Library.WriteErrorLog("Inventory Feed Service Started!");
        }
        private void timer1_Tick(object sender, ElapsedEventArgs e)
        {
            IFSReportingContext db_local = new IFSReportingContext(); //used for updating always the program
            //string message="";
           // var responsemessage = "";
            var responsestatus = "";
            var responsefilepath = "";
           string result = "";
            var timenowstring = DateTime.Now.ToString("H:mm:ss");
            TimeSpan timenow = TimeSpan.Parse(timenowstring, System.Globalization.CultureInfo.CurrentCulture);

            var datenow = DateTime.Now;
            var daynow = datenow.Date;
            var dayname = DateTime.Now.DayOfWeek.ToString().Substring(0,3); //Non, Tue
           

            try
            {
                var InventoryFeedList = (from t1 in db.tblInventoryFeeds
                                        join t2 in db.tblInventoryFeedProcesses on t1.if_id equals t2.if_id
                                         select new
                                         {
                                             t1.if_id,
                                             t2.ifp_id,
                                             t1.customer_no,
                                             t1.filetype_requested,
                                             t1.send_protocol,
                                             t1.protocol_address,
                                             t1.sendtime,
                                             t2.time_split,
                                             t1.sendbuyers_partno,
                                             t1.sendaaid_instead_brand_id,
                                             t1.sendday,
                                             t2.status,
                                             t2.datetime_updated,
                                             t1.fields,
                                             t2.current_pr
                                         } 
                                         ).Where(m => ( m.time_split <= timenow) && (m.status == "0") && (m.sendday.Contains(dayname))  )
                                         .Take(1).ToList();


               
                //Resetting status to zero
                var some = db_local.tblInventoryFeedProcesses.Where(x => EntityFunctions.TruncateTime(x.datetime_updated) != EntityFunctions.TruncateTime(DateTime.Now)).ToList();
                some.ForEach(a => {
                    a.status = "0";
                    a.datetime_updated = DateTime.Now;
                }
                 );
                db_local.SaveChanges();
                /////////

                //check if have current process else do not continue
                var current_process = (from cp in db.tblInventoryFeedProcesses
                                      select cp).Where(m=>m.current_pr==1).ToList();

                if (InventoryFeedList.Count > 0 && current_process.Count <= 0)
                {
                    foreach (var i in InventoryFeedList)
                    {
                       db_local = new IFSReportingContext();

                        var inv = new tblInventoryFeedProcess() //selecting for update
                        {
                            ifp_id = i.ifp_id,
                            status = "0"
                        };

                        db_local.tblInventoryFeedProcesses.Attach(inv);                  
                        inv.status = "1";
                        inv.current_pr = 1;
                        //update to 1
                        
                        db_local.SaveChanges();
                                               
                       result= DataCSVXLS(i.customer_no, i.filetype_requested, i.fields, db_local, inv);
                   
                       dynamic obj = Library.Json_Des(result);
                       responsestatus = obj["status"];
 
                        if (responsestatus == "success") //after creating file
                        {
                            responsefilepath = obj["filePath"];
                            if (i.send_protocol == "email")
                                result = Emails(i.customer_no, responsefilepath, i.protocol_address, db_local, inv);
                            else if (i.send_protocol == "ftp")
                                result = FTPs(i.customer_no, responsefilepath,i.protocol_address, db_local, inv);
                        }

                    }
                }
                else
                {
                   // Library.WriteErrorLog("No Feed");
                }
            }
            catch (Exception ex)
            {

                Library.WriteErrorLog(ex.Message + ":Main");
            }

         
            
        }


        protected override void OnStop()
        {
            timer1.Enabled = false;
            Library.WriteErrorLog("Inventory Feed Service stopped.");
        }
    }
}
