using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Net.Mail;

namespace TEMAutomatedReports
{
    class SendEmail
    {
        //jonathan.briggs@sentel.co.uk 20190611
        private static string _strmailClient = "pro.turbo-smtp.com";
        //private static string _strmailClient = "192.168.6.2";

        public static string MailClient
        {
            get { return _strmailClient; }
            set { _strmailClient = value; }
        }
        private static DataClassesDataContext _dataContext;
        public static DataClassesDataContext DbContext
        {
            get
            {
                if (_dataContext == null)
                {
                    _dataContext = new DataClassesDataContext();
                }
                return _dataContext;
            }
        }

        public static void SendEmails()
        {

        }

        public static void SendErrormail()
        {

        }

        public static void ErrorLog()
        {

        }

        public static void SendErrormessage()
        {


            //create the email message
            MailMessage mmAutomatedMessage = new MailMessage();

            //create a reply to address
            mmAutomatedMessage.ReplyTo = new MailAddress("servicedesk@sentel.co.uk");

            //set the priority of the mail message to high to make sure it goes out at quickest possible speed
            mmAutomatedMessage.Priority = MailPriority.High;

            //put a read receipt on each report. Can monitor who is looking at their reports
            mmAutomatedMessage.Headers.Add("Disposition-Notification-To", "ProAutomatedReports@sentel.co.uk");


            mmAutomatedMessage.From = new MailAddress("ProAutomatedReports@sentel.co.uk");

            //jonathan.briggs@sentel.co.uk 20190611
            string sendEmailsFrom = "roy.doherty@sentel.co.uk";
            string sendEmailsFromPassword = "QBURUS5N";
            //string sendEmailsFrom = "ProAutomatedReports@sentel.co.uk";
            //string sendEmailsFromPassword = "PR1pr2pr3";
            NetworkCredential cred = new NetworkCredential(sendEmailsFrom, sendEmailsFromPassword);
            mmAutomatedMessage.Subject = "Pro Error Log";
            //create the plain text version of the email
            IEnumerable<tbl_Reportstatus> rep = DbContext.tbl_Reportstatus.AsEnumerable();
            StringBuilder strBodyText = new StringBuilder();
            strBodyText.Append("<b> Dear Service team, </b>" +

                   "<br/>There are few pro reports failed today, please see below for more details");
            foreach (tbl_Reportstatus r in rep)
            {
                strBodyText.Append("<br/>Report id = " + r.schedule_id_PK + "--" + "Reason = " + r.Reportstatus_Status + "--" + "DateTime=" + r.Reportstatus_Date);
            }

            string strMediaType = "text/plain";

            //create an alternative view
            AlternateView avPlainText = AlternateView.CreateAlternateViewFromString(strBodyText.ToString(), null, strMediaType);

            // strBodyText = "";

            //create the media type for the html
            strMediaType = "text/html";

            //create an alternative view
            AlternateView avHTML = AlternateView.CreateAlternateViewFromString(strBodyText.ToString(), null, strMediaType);

            //add both views to the collection
            mmAutomatedMessage.AlternateViews.Add(avPlainText);
            mmAutomatedMessage.AlternateViews.Add(avHTML);

            //Attache the report/error log
            // mmAutomatedMessage.Attachments.Add(attach);

            //call the smtpclient to send the message - set to mail, will need changed to mail.sentel.co.uk when going live
            SmtpClient client = new SmtpClient(MailClient);

            try
            {
                client.Credentials = cred;
                client.Send(mmAutomatedMessage);
            }
            catch
            {

            }

        }

        public static void InsertPortfoliodetails(string filename, string strEmailAddress, int PFID, int Reportid)
        {

            string[] email = strEmailAddress.Split(',');

            foreach (string s in email)
            {
                if (s.Contains('@'))
                {
                    var pfreports = new tbl_PortfolioReport();
                    pfreports.PortfolioReport_ReportName = filename;
                    pfreports.PortfolioReport_Email = s;
                    pfreports.Protfolio_ID = PFID;
                    pfreports.schedule_id_PK = Reportid;
                    pfreports.PortfolioReport_Status = false;
                    DbContext.tbl_PortfolioReports.InsertOnSubmit(pfreports);
                }
            }
            try
            {
                DbContext.SubmitChanges();
            }
            catch
            { }




        }
        public static void sendEmails()
        {
            List<string> Emailreports;
            // Normal reports
            var portfolio = from reports in DbContext.tbl_PortfolioReports
                            where reports.Protfolio_ID == 0
                            && reports.PortfolioReport_Status == false

                            group reports by new { reports.PortfolioReport_Email } into groupclause

                            select new
                            {
                                Email = groupclause.Key.PortfolioReport_Email,
                                Reportname = groupclause.Select(s => s.PortfolioReport_ReportName),
                                ReportID = groupclause.Select(s => s.schedule_id_PK ?? 0)

                            };


            foreach (var report in portfolio)
            {
                Emailreports = new List<string>();
                foreach (var v in report.Reportname)
                {
                    Emailreports.Add(v.ToString());
                }
                try
                {
                    send(report.Email.ToString(), Emailreports, report.ReportID);
                }
                catch { }
            }

            // Portfolio reorts
            var custportfolio = from reports in DbContext.tbl_PortfolioReports
                                where reports.Protfolio_ID != 0
                                && reports.PortfolioReport_Status == false
                                group reports by new { reports.PortfolioReport_Email, reports.Protfolio_ID } into groupclause

                                select new
                                {
                                    Email = groupclause.Key.PortfolioReport_Email,
                                    Reportname = groupclause.Select(s => s.PortfolioReport_ReportName),
                                    PFID = groupclause.Key.Protfolio_ID ?? 0,
                                    ReportID = groupclause.Select(s => s.schedule_id_PK ?? 0)

                                    // Name  =  groupclause.Where(s=>s.tbl_Portfoilolink.Protfolio_ID == groupclause.Key.Protfolio_ID).Select(s=>s.tbl_Portfoilolink.Protfolio_Name) 
                                };

            foreach (var report in custportfolio)
            {
                Emailreports = new List<string>();
                foreach (var v in report.Reportname)
                {
                    Emailreports.Add(v.ToString());
                }
                try
                {
                    send(report.Email.ToString(), Emailreports, report.PFID, report.ReportID);
                }
                catch { }
            }
            var reportssent = DbContext.tbl_PortfolioReports.Where(s => s.PortfolioReport_Status == true);
            DbContext.tbl_PortfolioReports.DeleteAllOnSubmit(reportssent);
            DbContext.SubmitChanges();

        }

        public static void send(string Email, List<string> report, IEnumerable<int> reportid)
        {


            //create the email message
            MailMessage mmAutomatedMessage = new MailMessage();

            //create a reply to address
            mmAutomatedMessage.ReplyTo = new MailAddress("servicedesk@sentel.co.uk");

            //set the priority of the mail message to high to make sure it goes out at quickest possible speed
            mmAutomatedMessage.Priority = MailPriority.High;

            //put a read receipt on each report. Can monitor who is looking at their reports
            mmAutomatedMessage.Headers.Add("Disposition-Notification-To", "ProAutomatedReports@sentel.co.uk");

            //create From address
            mmAutomatedMessage.From = new MailAddress("ProAutomatedReports@sentel.co.uk");


            //jonathan.briggs@sentel.co.uk 20190611
            string sendEmailsFrom = "roy.doherty@sentel.co.uk";
            string sendEmailsFromPassword = "QBURUS5N";
            //string sendEmailsFrom = "ProAutomatedReports@sentel.co.uk";
            //string sendEmailsFromPassword = "PR1pr2pr3";
            NetworkCredential cred = new NetworkCredential(sendEmailsFrom, sendEmailsFromPassword);
            mmAutomatedMessage.To.Add(Email);
            //create subject
            mmAutomatedMessage.Subject = "Pro Automated Report";

            //create the plain text version of the email
            string strBodyText = @"Dear Customer, 

            Please open all image attachments prior to the HTML report in order to cache the charts.

            Please find enclosed your automated report as requested:

            Many Thanks,
            Pro Team 
            Sentel Indepedant Ltd 
            15 Mckibbin House 
            Eastbank Road 
            Carryduff 
            Belfast 
            BT8 8BD 

            T: +44 (0)28 9081 5555 
            F: +44 (0)28 9081 1055 
            E: servicedesk@sentel.co.uk  

            Keep up to date on http://www.sentel.co.uk 

            Follow us on twitter at http://twitter.com/Sentel_Ind 

            Follow us on facebook at http://www.facebook.com/pages/Belfast/Sentel-Independent/249597243565";

            //create the media type for the plain text
            string strMediaType = "text/plain";

            //create an alternative view
            AlternateView avPlainText = AlternateView.CreateAlternateViewFromString(strBodyText, null, strMediaType);

            //create the html version of the email
            strBodyText = @"<font face ='verdana' size='3'><p>Dear <b><i>Customer</i></b>,</p>
            
            

            <p>Please find enclosed your automated report as requested:</p>

            <p>Many Thanks,</p>
            <p><b>Pro Team</b></p>
            <p>Sentel 
            <br />15 McKibbin House
            <br />Eastbank Road
            <br />Carryduff
            <br />Belfast
            <br />BT8 8BD
            </p>
            <p><font face ='verdana' size='2'>T: +44 (0)28 9081 5555
            <br />F: +44 (0)28 9081 1055
            <br />E: servicedesk@sentel.co.uk 
            </font>
            </p></font>
            <font face ='verdana' size='1'>
            <p>
            <a href='http://www.sentel.co.uk'>www.Sentel.co.uk</a>
            <br /><img src='http://www.sentelcallmanagerpro.com//images//SentelLogo.png' alt='Sentel' title='Sentel' />
            </p>
            <p>
            Follow Sentel on Twitter
            <br /><img src='http://www.sentelcallmanagerpro.com//images//twitter.png' alt='Twitter' title='Twitter' />
            <br />Ask us a question or keep up to date with our latest business news and events using Twitter
            <br /><a href='http://twitter.com/Sentel_Ind'>Sentel on Twitter</a>
            </p>
            <p>
            Follow Sentel on Facebook
            <br /><img src='http://www.sentelcallmanagerpro.com//images//facebook.png' alt='Facebook' title='FaceBook' />
            <br />Add us to get all the latest upgrade news and developments on our products using Facebook
            <br /><a href='http://www.facebook.com/pages/Belfast/Sentel-Independent/249597243565'>Sentel on Facebook</a>
            </p>
            </font>";

            //create the media type for the html
            strMediaType = "text/html";

            //create an alternative view
            AlternateView avHTML = AlternateView.CreateAlternateViewFromString(strBodyText, null, strMediaType);

            //add both views to the collection
            mmAutomatedMessage.AlternateViews.Add(avPlainText);
            mmAutomatedMessage.AlternateViews.Add(avHTML);
            try
            {
                //Attache the report/error log
                Attachment attach;
                for (int i = 0; i < report.Count(); i++)
                {
                    attach = new Attachment(report[i]);

                    mmAutomatedMessage.Attachments.Add(attach);
                }


                SmtpClient client = new SmtpClient(MailClient);




                client.Timeout = 40000;
                client.Credentials = cred;
                client.Send(mmAutomatedMessage);
                DbContext.tbl_PortfolioReports.Where(s => s.PortfolioReport_Email == Email).ToList().ForEach(s => s.PortfolioReport_Status = true);
                DbContext.SubmitChanges();


            }
            catch (Exception ex)
            {
                foreach (int v in reportid)
                {
                    GenerateReports.ReportStatus(v, "ErrEmail");

                }
            }

            //ProgressText.Append("Email sent to '" + Email + "'. Completed \r\n");
            //tbProgress.Text = ProgressText.ToString(); ;

        }
        public static void send(string Email, List<string> report, int id, IEnumerable<int> reportid)
        {


            //create the email message
            MailMessage mmAutomatedMessage = new MailMessage();

            //create a reply to address
            mmAutomatedMessage.ReplyTo = new MailAddress("servicedesk@sentel.co.uk");

            //set the priority of the mail message to high to make sure it goes out at quickest possible speed
            mmAutomatedMessage.Priority = MailPriority.High;

            //put a read receipt on each report. Can monitor who is looking at their reports
            mmAutomatedMessage.Headers.Add("Disposition-Notification-To", "ProAutomatedReports@sentel.co.uk");

            //create From address
            mmAutomatedMessage.From = new MailAddress("ProAutomatedReports@sentel.co.uk");
            //jonathan.briggs@sentel.co.uk 20190611
            string sendEmailsFrom = "roy.doherty@sentel.co.uk";
            string sendEmailsFromPassword = "QBURUS5N";
            //string sendEmailsFrom = "ProAutomatedReports@sentel.co.uk";
            //string sendEmailsFromPassword = "PR1pr2pr3";
            NetworkCredential cred = new NetworkCredential(sendEmailsFrom, sendEmailsFromPassword);
            mmAutomatedMessage.To.Add(Email);
            //create subject
            mmAutomatedMessage.Subject = "Pro Automated Report" + " - " + DbContext.tbl_Portfoilolinks.Where(s => s.Protfolio_ID == id).Select(s => s.Protfolio_Name ?? "").SingleOrDefault();

            //create the plain text version of the email
            string strBodyText = @"Dear Customer, 

            Please open all image attachments prior to the HTML report in order to cache the charts.

            Please find enclosed your automated report as requested:

            Many Thanks,
            Pro Team 
            Sentel Indepedant Ltd 
            15 Mckibbin House 
            Eastbank Road 
            Carryduff 
            Belfast 
            BT8 8BD 

            T: +44 (0)28 9081 5555 
            F: +44 (0)28 9081 1055 
            E: servicedesk@sentel.co.uk  

            Keep up to date on http://www.sentel.co.uk 

            Follow us on twitter at http://twitter.com/Sentel_Ind 

            Follow us on facebook at http://www.facebook.com/pages/Belfast/Sentel-Independent/249597243565";

            //create the media type for the plain text
            string strMediaType = "text/plain";

            //create an alternative view
            AlternateView avPlainText = AlternateView.CreateAlternateViewFromString(strBodyText, null, strMediaType);

            //create the html version of the email
            strBodyText = @"<font face ='verdana' size='3'><p>Dear <b><i>Customer</i></b>,</p>
            
            

            <p>Please find enclosed your automated report as requested:</p>

            <p>Many Thanks,</p>
            <p><b>Pro Team</b></p>
            <p>Sentel 
            <br />15 McKibbin House
            <br />Eastbank Road
            <br />Carryduff
            <br />Belfast
            <br />BT8 8BD
            </p>
            <p><font face ='verdana' size='2'>T: +44 (0)28 9081 5555
            <br />F: +44 (0)28 9081 1055
            <br />E: servicedesk@sentel.co.uk 
            </font>
            </p></font>
            <font face ='verdana' size='1'>
            <p>
            <a href='http://www.sentel.co.uk'>www.Sentel.co.uk</a>
            <br /><img src='http://www.sentelcallmanagerpro.com//images//SentelLogo.png' alt='Sentel' title='Sentel' />
            </p>
            <p>
            Follow Sentel on Twitter
            <br /><img src='http://www.sentelcallmanagerpro.com//images//twitter.png' alt='Twitter' title='Twitter' />
            <br />Ask us a question or keep up to date with our latest business news and events using Twitter
            <br /><a href='http://twitter.com/Sentel_Ind'>Sentel on Twitter</a>
            </p>
            <p>
            Follow Sentel on Facebook
            <br /><img src='http://www.sentelcallmanagerpro.com//images//facebook.png' alt='Facebook' title='FaceBook' />
            <br />Add us to get all the latest upgrade news and developments on our products using Facebook
            <br /><a href='http://www.facebook.com/pages/Belfast/Sentel-Independent/249597243565'>Sentel on Facebook</a>
            </p>
            </font>";

            //create the media type for the html
            strMediaType = "text/html";

            //create an alternative view
            AlternateView avHTML = AlternateView.CreateAlternateViewFromString(strBodyText, null, strMediaType);

            //add both views to the collection
            mmAutomatedMessage.AlternateViews.Add(avPlainText);
            mmAutomatedMessage.AlternateViews.Add(avHTML);

            //Attache the report/error log
            try
            {

                Attachment attach;
                for (int i = 0; i < report.Count(); i++)
                {
                    attach = new Attachment(report[i]);

                    mmAutomatedMessage.Attachments.Add(attach);
                }


                SmtpClient client = new SmtpClient(MailClient);



                client.Timeout = 40000;
                client.Credentials = cred;
                client.Send(mmAutomatedMessage);
                DbContext.tbl_PortfolioReports.Where(s => s.PortfolioReport_Email == Email).ToList().ForEach(s => s.PortfolioReport_Status = true);
                DbContext.SubmitChanges();
            }
            catch
            {

                foreach (int v in reportid)
                {
                    GenerateReports.ReportStatus(v, "ErrEmail");
                }
            }

            //ProgressText.Append("Email sent to '" + Email + "'. Completed \r\n");
            //tbProgress.Text = ProgressText.ToString(); ;

        }
    }
}
