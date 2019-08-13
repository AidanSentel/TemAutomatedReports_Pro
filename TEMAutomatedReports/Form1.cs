using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Configuration;
using System.Net.Mail;
using System.Collections;
using System.IO;
using System.Reflection;
using Dundas.Charting.WinControl;
using Doc = MigraDoc.DocumentObjectModel;
using Shapes = MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.Rendering;
using MigraDoc.RtfRendering;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes.Charts;
using iTextSharp.text;
using iTextSharp.text.pdf;




/*
 TEMAutomated Reports is the TEM Scheduler.
 The code runs a sp which puts all reports for that day into a working folder.
 A further sp returns all the data needed for a the report.
 This report is then generated along with a message file.
 
 
 Need to check if datapump has run for the site otherwise we should not send the schedule as it will be blank
 * 
 * Phani -NOV 2013 -- code changes-----
 * All code is now replaced with new  classes and linq to sql
 * imlemented deparmental break down reports
 * added quaterly and yearly schedule report option
 * html report is modified
 * ids are implemented at the end of report file name
 * totals are now alligned at the buttom of the column
 * there is a functionality for xml report but not fully imlemented
 * improved the performance by 70%.
 * rerun facility is created for service team...
 * ----------- end----------------------------------------------
 */


namespace TEMAutomatedReports
{
    public partial class Form1 : Form
    {
        #region Properties -----------------------------------------------------------------------------------------------------------

        private string _strConnectionString = ConfigurationManager.AppSettings["TEMConnectionString"];
        private bool _blnReportHasData = false;
        private StringBuilder _sbProgress = new StringBuilder();
        private StringBuilder _sbError = new StringBuilder();

        public string TEMConnectionString
        {
            get { return _strConnectionString; }
            set { _strConnectionString = value; }
        }
        public bool ReportHasData
        {
            get { return _blnReportHasData; }
            set { _blnReportHasData = value; }
        }

        public StringBuilder ProgressText
        {
            get { return _sbProgress; }
            set { _sbProgress = value; }
        }

        public StringBuilder ErrorText
        {
            get { return _sbError; }
            set { _sbError = value; }
        }

        private string _strmailClient = "192.168.6.2";
        public string MailClient
        {
            get { return _strmailClient; }
            set { _strmailClient = value; }
        }
        private int _strlevel = 0;
        public int Level
        {
            get { return _strlevel; }
            set { _strlevel = value; }
        }
        private int _Callcount = 0;
        public int Callcount
        {
            get { return _Callcount; }
            set { _Callcount = value; }
        }

        private DataClassesDataContext _dataContext;
        protected DataClassesDataContext DbContext
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

        #endregion --------------------------------------------------------------------------------------------------------------------

        #region Functions & Methods --------------------------------------------------------------------------------------------------
        /*
        public void CreateDailySchedules()
        {
            SqlConnection oConn = new SqlConnection(TEMConnectionString);

            oConn.Open();

            //stored procedure firstly clears down the processreports table to clean up for the day before
            //and creates entries for all reports due to be sent that day based on date selections
            SqlCommand oCmd = new SqlCommand("spCreateDailySchedules", oConn);
            oCmd.CommandType = CommandType.StoredProcedure;
            oCmd.ExecuteNonQuery();

            oConn.Close();
        }

        public void Updatefilters(string ListofFilters, int ID1)
        {

            SqlConnection oConn = new SqlConnection(TEMConnectionString);

            oConn.Open();
            SqlParameter p1;

            SqlCommand oCmd = new SqlCommand("spupdatefilterinautometedreport", oConn);


            oCmd.CommandType = CommandType.StoredProcedure;

            p1 = new SqlParameter("@filter", SqlDbType.VarChar, 2000);
            p1.Value = ListofFilters;
            oCmd.Parameters.Add(p1);

            p1 = new SqlParameter("@ID", SqlDbType.Int);
            p1.Value = ID1;
            oCmd.Parameters.Add(p1);
            oCmd.ExecuteNonQuery();
            oConn.Close();

        }

        public void DeleteProcessedReports()
        {
            SqlConnection oConn = new SqlConnection(TEMConnectionString);

            oConn.Open();

            SqlCommand oCmd = new SqlCommand("spDeleteProcessedReports", oConn);
            oCmd.CommandType = CommandType.StoredProcedure;
            oCmd.ExecuteNonQuery();

            oConn.Close();
        }

        public void sendEmails()
        {
            List<string> Emailreports;
            var portfolio = from reports in DbContext.tbl_PortfolioReports.AsEnumerable()

                            group reports by new { reports.PortfolioReport_Email } into groupclause

                            select new
                            {
                                Email = groupclause.Key.PortfolioReport_Email,
                                Reportname = groupclause.Select(s => s.PortfolioReport_ReportName)
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
                    send(report.Email.ToString(), Emailreports);
                }
                catch { }
            }



            DbContext.ExecuteCommand("Truncate Table tbl_PortfolioReport ");
        }
        public void sendportfolioreports()
        {

            List<string> emailslist = new List<string>();
            //var Portfolio = DbContext.tbl_customportfolioreports.Where(s=>strUniqueDateTime.Substring(0,8).Contains(s.Customportfolio_Link)).Select(s=>s.schedule_id_PK);
            var Portfolio = from db in DbContext.tbl_customportfolioreports.AsEnumerable()
                            where db.Customportfolio_Run == true
                            group db by new { db.Customportfolio_Link } into groupclause
                            select new
                            {
                                ID = groupclause.Key.Customportfolio_Link,
                                ScheduleID = groupclause.Select(s => s.schedule_id_PK),
                                Reportname = groupclause.Select(s => s.Customportfolio_Report),
                                Email = DbContext.tbl_automatedreports_tests.Where(s => groupclause.Select(r => r.schedule_id_PK).Contains(s.schedule_id_PK)).Select(s => s.schedule_emailaddresses)


                            };
            foreach (var g in Portfolio)
            {

                var email = g.Email.Distinct();
                foreach (var e in email)
                {
                    string[] em = e.Split(',');
                    for (int i = 0; i < em.Count(); i++)
                    {
                        if (em[i] != "")
                        {
                            if (!emailslist.Contains(em[i]))
                            {
                                emailslist.Add(em[i]);

                            }
                        }

                    }

                }


                sendportfolio(emailslist, g.Reportname);

            }
        }
        #region "Generating report"
        public ArrayList GenerateCDRResults(string strStoredProcedure, string strListofFilters, string strEmailAddresses, string strColumns, string Chosennodelist, int id1, string Choosenname, string totals)
        {
            string[] choosennode = Chosennodelist.Split('=');
            string param = "'" + choosennode[1] + "'";
            string name = choosennode[0] + "=";

            Chosennodelist = name + param;

            SqlConnection oConn = new SqlConnection(TEMConnectionString);

            oConn.Open();


            SqlCommand oCmdCDR = new SqlCommand("exec " + strStoredProcedure + " " + strListofFilters + "," + Chosennodelist, oConn);
            oCmdCDR.CommandTimeout = 3000;
            //oConn.ConnectionTimeout = 3000;
            ArrayList alExport = new ArrayList();


            try
            {
                SqlDataReader drReaderCDR = oCmdCDR.ExecuteReader();

                //drReaderCDR.Read();


                //infeed_startdatetime@15,infeed_firstextension,Name,infeed_originatortrunkgroup,infeed_secondextension,infeed_ringduration,ring_duration,infeed_connectduration,connected_duration,infeed_callingnumber,infeed_callednumber
                string[] Columns = strColumns.Split(',');

                //15,5,8,79,46,
                //string [] Length = strColumns.Split('@')


                foreach (string Column in Columns)
                {
                    string[] arrColumnNameAndLength = Column.Split('@', '#');

                    //do function and pass in column name and amount
                    alExport.Add(arrColumnNameAndLength[0]);
                }

                alExport.Add("\r\n");

                // if (drReaderCDR.HasRows == true)
                // {
                ReportHasData = false;
                Callcount = 0;
                while (drReaderCDR.Read())
                {
                    ReportHasData = true;
                    // needs changed to dynamically know the table column etc - 

                    foreach (string Column in Columns)
                    {
                        string[] arrColumnNameAndLength = Column.Split('@', '#');
                        string strResult = drReaderCDR[arrColumnNameAndLength[0]].ToString();
                        alExport.Add(ColumnValue(strResult, Int32.Parse(arrColumnNameAndLength[1]), arrColumnNameAndLength[2]));
                    }
                    Callcount++;
                    alExport.Add("\r\n");

                }
                drReaderCDR.Close();
                //  }
                //  else
                //  {
                //      ReportHasData = false;
                //      tbError.Text = "Email to '" + strEmailAddresses + "' could not be sent. No data available for the report";
                //  }

            }
            catch
            {
               // tbError.Text = "Problem connecting to the database.";
            }


            finally
            {


                oConn.Close();

            }
            return alExport;
        }
        private ArrayList GetReportdata(string strStoredProcedure, string strListofFilters, string strColumns, string Chosennodelist)
        {
            string[] choosennode = Chosennodelist.Split('=');
            string param = "'" + choosennode[1] + "'";
            string name = choosennode[0] + "=";

            Chosennodelist = name + param;
            DataSet dsCDRResults = new DataSet();
            SqlConnection oConn = new SqlConnection(TEMConnectionString);
            SqlCommand oCmd = new SqlCommand("exec " + strStoredProcedure + " " + strListofFilters + "," + Chosennodelist, oConn);
            SqlDataAdapter sqlDA = new SqlDataAdapter(oCmd);
            ArrayList alExport = new ArrayList();
            oCmd.CommandTimeout = 3000;
            try
            {
                sqlDA.Fill(dsCDRResults);
            }
            catch
            { }
            if (dsCDRResults.Tables[0].Rows.Count > 0)
            {
                ReportHasData = true;
                string[] Columns = strColumns.Split(',');
                foreach (string Column in Columns)
                {
                    string[] arrColumnNameAndLength = Column.Split('@', '#');
                    alExport.Add(arrColumnNameAndLength[0]);
                }

                alExport.Add("\r\n");
                foreach (DataRow dr in dsCDRResults.Tables[0].Rows)
                {
                    foreach (DataColumn dc in dsCDRResults.Tables[0].Columns)
                    {
                        alExport.Add(Convert.ToString(dr[dc]) + "   ");
                    }
                    alExport.Add("\r\n");
                }
            }
            else
            {
                ReportHasData = false;
            }
            return alExport;




        }
        #endregion
        #region "sending email.."
        private void sendportfolio(List<string> emailslist, IEnumerable<string> iEnumerable)
        {
            List<string> report = new List<string>();

            foreach (var v in iEnumerable)
            {
                report.Add(v);

            }

            //create the email message
            MailMessage mmAutomatedMessage = new MailMessage();

            //create a reply to address
            mmAutomatedMessage.ReplyTo = new MailAddress("servicedesk@sentel.co.uk ");

            //set the priority of the mail message to high to make sure it goes out at quickest possible speed
            mmAutomatedMessage.Priority = MailPriority.High;

            //put a read receipt on each report. Can monitor who is looking at their reports
            mmAutomatedMessage.Headers.Add("Disposition-Notification-To", "ProAutomatedReports@sentel.co.uk");

            //create From address
            mmAutomatedMessage.From = new MailAddress("ProAutomatedReports@sentel.co.uk");
            string sendEmailsFrom = "ProAutomatedReports@sentel.co.uk";
            string sendEmailsFromPassword = "PR1pr2pr3";
            NetworkCredential cred = new NetworkCredential(sendEmailsFrom, sendEmailsFromPassword);

            //string strTrimStartEmail = strEmail.TrimStart(',');
            //string strTrimEndEmail = strTrimStartEmail.TrimEnd(',');

            //string[] strEmails = strTrimEndEmail.Split(',');

            // To do testing always use this name
            //string mail1 = "krishna@sentel.co.uk";
            //mmAutomatedMessage.To.Add(mail1);
            foreach (string Email in emailslist)
            {
                //create to address
                mmAutomatedMessage.To.Add(Email);
            }
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

            //Attache the report/error log
            Attachment attach;
            for (int i = 0; i < report.Count(); i++)
            {
                attach = new Attachment(report[i]);

                mmAutomatedMessage.Attachments.Add(attach);
            }

            //if (strExtension == ".html" & blnHaveChart == true)
            //{
            //    //Attachment attachimageheader = new Attachment("c:\\temp\\Header.png");
            //    //mmAutomatedMessage.Attachments.Add(attachimageheader);

            //        Attachment attachimagehtmlbar = new Attachment("Z:\\inetpub\\wwwroot\\Pro_Application\\images\\chart" + strDateTime + ".png");
            //        Attachment attachimagehtmlpie = new Attachment("Z:\\inetpub\\wwwroot\\Pro_Application\\images\\pie" + strDateTime + ".png");

            //        mmAutomatedMessage.Attachments.Add(attachimagehtmlbar);
            //        mmAutomatedMessage.Attachments.Add(attachimagehtmlpie);

            //}

            //call the smtpclient to send the message - set to mail, will need changed to mail.sentel.co.uk when going live
            SmtpClient client = new SmtpClient(MailClient);

            try
            {
                TimeSpan ts = new TimeSpan(0, 0, 2);
                System.Threading.Thread.Sleep(ts);
                client.Timeout = 400;
                client.Credentials = cred;
                client.Send(mmAutomatedMessage);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                //ErrorText.Append("There has been a problem connecting to the mail sever. \r\n");
                //tbError.Text = ErrorText.ToString();
            }
            // ProgressText.Append("Email sent to '" + Email + "'. Completed \r\n");
            //tbProgress.Text = ProgressText.ToString(); ;
        }

        public void send(string Email, List<string> report)
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
            string sendEmailsFrom = "ProAutomatedReports@sentel.co.uk";
            string sendEmailsFromPassword = "PR1pr2pr3";
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

            //Attache the report/error log
            Attachment attach;
            for (int i = 0; i < report.Count(); i++)
            {
                attach = new Attachment(report[i]);

                mmAutomatedMessage.Attachments.Add(attach);
            }


            SmtpClient client = new SmtpClient(MailClient);

            try
            {

                TimeSpan ts = new TimeSpan(0, 0, 2);
                System.Threading.Thread.Sleep(ts);
                client.Timeout = 40000;
                client.Credentials = cred;
                client.Send(mmAutomatedMessage);
            }
            catch
            {

                CreateErrorMessage(Email, 0, DateTime.Now, "ALL reports", "To this email address");
                // MessageBox.Show(ex.ToString());

                //  ErrorText.Append("There has been a problem connecting to the mail sever. \r\n");
                //  tbError.Text = ErrorText.ToString();
            }

            //ProgressText.Append("Email sent to '" + Email + "'. Completed \r\n");
            //tbProgress.Text = ProgressText.ToString(); ;

        }

        public void CreateMailMessage(string strEmail, string strReport, string strExtension, string strDateTime, bool blnHaveChart, string givenname)
        {


            Attachment attach = new Attachment("c:\\temp\\" + givenname + " - " + strReport + strDateTime + strExtension);

            //create the email message
            MailMessage mmAutomatedMessage = new MailMessage();

            //create a reply to address
            mmAutomatedMessage.ReplyTo = new MailAddress("servicedesk@sentel.co.uk ");

            //set the priority of the mail message to high to make sure it goes out at quickest possible speed
            mmAutomatedMessage.Priority = MailPriority.High;

            //put a read receipt on each report. Can monitor who is looking at their reports
            mmAutomatedMessage.Headers.Add("Disposition-Notification-To", "ProAutomatedReports@sentel.co.uk");

            //create From address
            mmAutomatedMessage.From = new MailAddress("ProAutomatedReports@sentel.co.uk");
            string sendEmailsFrom = "ProAutomatedReports@sentel.co.uk";
            string sendEmailsFromPassword = "PR1pr2pr3";
            NetworkCredential cred = new NetworkCredential(sendEmailsFrom, sendEmailsFromPassword);

            string strTrimStartEmail = strEmail.TrimStart(',');
            string strTrimEndEmail = strTrimStartEmail.TrimEnd(',');

            string[] strEmails = strTrimEndEmail.Split(',');

            // To do testing always use this name
            //string mail1 = "krishna@sentel.co.uk";
            //mmAutomatedMessage.To.Add(mail1);

            foreach (string Email in strEmails)
            {
                //create to address
                mmAutomatedMessage.To.Add(Email);
            }
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
            <p>Sentel Independent Ltd
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
            mmAutomatedMessage.Attachments.Add(attach);

            //if (strExtension == ".html" & blnHaveChart == true)
            //{
            //    //Attachment attachimageheader = new Attachment("c:\\temp\\Header.png");
            //    //mmAutomatedMessage.Attachments.Add(attachimageheader);

            //        Attachment attachimagehtmlbar = new Attachment("Z:\\inetpub\\wwwroot\\Pro_Application\\images\\chart" + strDateTime + ".png");
            //        Attachment attachimagehtmlpie = new Attachment("Z:\\inetpub\\wwwroot\\Pro_Application\\images\\pie" + strDateTime + ".png");

            //        mmAutomatedMessage.Attachments.Add(attachimagehtmlbar);
            //        mmAutomatedMessage.Attachments.Add(attachimagehtmlpie);

            //}

            //call the smtpclient to send the message - set to mail, will need changed to mail.sentel.co.uk when going live
            SmtpClient client = new SmtpClient(MailClient);

            try
            {
                // client.Credentials = cred;
                client.Send(mmAutomatedMessage);

            }
            catch
            {
                // ErrorText.Append("There has been a problem connecting to the mail sever. \r\n");
                //  tbError.Text = ErrorText.ToString();
            }
            // ProgressText.Append("Email sent to '" + strEmail + "'. Completed \r\n");
            //tbProgress.Text = ProgressText.ToString(); ;
        }

        public void CreateMailMessage(string strEmail, string strReport, string strExtension, string strDateTime, string givenname)
        {
            //strEmail = ",Adam.Shilliday@sentel.co.uk,";
            Attachment attach = new Attachment("c:\\temp\\" + givenname + " - " + strReport + strDateTime + strExtension);

            //create the email message
            MailMessage mmAutomatedMessage = new MailMessage();

            //create a reply to address
            mmAutomatedMessage.ReplyTo = new MailAddress("servicedesk@sentel.co.uk ");

            //set the priority of the mail message to high to make sure it goes out at quickest possible speed
            mmAutomatedMessage.Priority = MailPriority.High;

            //put a read receipt on each report. Can monitor who is looking at their reports
            mmAutomatedMessage.Headers.Add("Disposition-Notification-To", "glen.adamson@sentel.co.uk");

            //create From address
            mmAutomatedMessage.From = new MailAddress("ProAutomatedReports@sentel.co.uk");
            string sendEmailsFrom = "ProAutomatedReports@sentel.co.uk";
            string sendEmailsFromPassword = "PR1pr2pr3";
            NetworkCredential cred = new NetworkCredential(sendEmailsFrom, sendEmailsFromPassword);


            string strTrimStartEmail = strEmail.TrimStart(',');
            string strTrimEndEmail = strTrimStartEmail.TrimEnd(',');

            string[] strEmails = strTrimEndEmail.Split(',');
            // string mail1 = "krishna@sentel.co.uk";
            // mmAutomatedMessage.To.Add(mail1);

            foreach (string Email in strEmails)
            {
                // create to address
                mmAutomatedMessage.To.Add(Email);
            }
            //create subject
            mmAutomatedMessage.Subject = "Pro Automated Report";

            //create the plain text version of the email
            string strBodyText = @"Dear Customer, 

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
            <p>Sentel Independent Ltd
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
            <br /><img src='http://www.sentelcallmanagerpro.com//images//facebook.jpg' alt='Facebook' title='FaceBook' />
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
            mmAutomatedMessage.Attachments.Add(attach);

            if (strExtension == ".html")
            {
                //Attachment attachimageheader = new Attachment("c:\\temp\\Header.png");
                //mmAutomatedMessage.Attachments.Add(attachimageheader);

                Attachment attachimagehtmlbar = new Attachment("c:\\visualstudio\\projects\\TemAutomatedReports\\images\\chart.png");
                Attachment attachimagehtmlpie = new Attachment("c:\\visualstudio\\projects\\TemAutomatedReports\\images\\pie.png");
                mmAutomatedMessage.Attachments.Add(attachimagehtmlbar);
                mmAutomatedMessage.Attachments.Add(attachimagehtmlpie);
            }

            //call the smtpclient to send the message - set to mail, will need changed to mail.sentel.co.uk when going live
            SmtpClient client = new SmtpClient(MailClient);

            try
            {
                //client.Credentials = cred;
                client.Send(mmAutomatedMessage);

            }
            catch
            {
                //  ErrorText.Append("There has been a problem connecting to the mail sever. \r\n");
                //  tbError.Text = ErrorText.ToString();
            }
            //ProgressText.Append("Email sent to '" + strEmail + "'. Completed \r\n");
            //tbProgress.Text = ProgressText.ToString(); ;
        }

        public void CreateMailMessage(string strEmail, string strFilename)
        {
            Attachment attach = new Attachment(strFilename);

            //create the email message
            MailMessage mmAutomatedMessage = new MailMessage();

            //create a reply to address
            mmAutomatedMessage.ReplyTo = new MailAddress("servicedesk@sentel.co.uk ");

            //set the priority of the mail message to high to make sure it goes out at quickest possible speed
            mmAutomatedMessage.Priority = MailPriority.High;

            //put a read receipt on each report. Can monitor who is looking at their reports
            mmAutomatedMessage.Headers.Add("Disposition-Notification-To", "ProAutomatedReports@sentel.co.uk");

            //create From address
            mmAutomatedMessage.From = new MailAddress("ProAutomatedReports@sentel.co.uk");

            string strTrimStartEmail = strEmail.TrimStart(',');
            string strTrimEndEmail = strTrimStartEmail.TrimEnd(',');

            string[] strEmails = strTrimEndEmail.Split(',');

            //string mail1 = "krishna@sentel.co.uk";
            //mmAutomatedMessage.To.Add(mail1);
            foreach (string Email in strEmails)
            {
                //create to address
                mmAutomatedMessage.To.Add(Email);
            }

            string sendEmailsFrom = "ProAutomatedReports@sentel.co.uk";
            string sendEmailsFromPassword = "PR1pr2pr3";
            NetworkCredential cred = new NetworkCredential(sendEmailsFrom, sendEmailsFromPassword);
            //create to address
            //mmAutomatedMessage.To.Add("krishna@sentel.co.uk");

            //create subject
            mmAutomatedMessage.Subject = "Pro Automated Report";

            //create the plain text version of the email
            string strBodyText = @"Dear Customer, 

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
            <p>Sentel Independent Ltd
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
            mmAutomatedMessage.Attachments.Add(attach);

            //call the smtpclient to send the message - set to mail, will need changed to mail.sentel.co.uk when going live
            SmtpClient client = new SmtpClient(MailClient);

            try
            {
                // client.Credentials = cred;
                client.Send(mmAutomatedMessage);

            }
            catch
            {
                //   ErrorText.Append("There has been a problem connecting to the mail sever. \r\n");
                //    tbError.Text = ErrorText.ToString();
            }
            // ProgressText.Append("Email sent to '" + strEmail + "'. Completed \r\n");
            // tbProgress.Text = ProgressText.ToString(); ;
        }

        public void CreateErrorMessage(string strEmail, int id1, DateTime Datecreated, string Reportname)
        {

            //Attachment attach = new Attachment("c:\\temp\\" + "TemautomatedReportslogfile"+ id1 + Reportname + ".txt");

            //create the email message
            MailMessage mmAutomatedMessage = new MailMessage();

            //create a reply to address
            mmAutomatedMessage.ReplyTo = new MailAddress("servicedesk@sentel.co.uk ");

            //set the priority of the mail message to high to make sure it goes out at quickest possible speed
            mmAutomatedMessage.Priority = MailPriority.High;

            //put a read receipt on each report. Can monitor who is looking at their reports
            mmAutomatedMessage.Headers.Add("Disposition-Notification-To", "ProAutomatedReports@sentel.co.uk");

            //create From address
            //mmAutomatedMessage.From = new MailAddress("TEMAutomatedReports@sentel.co.uk");

            //create to address
            //mmAutomatedMessage.To.Add("Glen.adamson@sentel.co.uk");
            mmAutomatedMessage.From = new MailAddress("ProAutomatedReports@sentel.co.uk");

            string sendEmailsFrom = "ProAutomatedReports@sentel.co.uk";
            string sendEmailsFromPassword = "PR1pr2pr3";
            NetworkCredential cred = new NetworkCredential(sendEmailsFrom, sendEmailsFromPassword);

            //string strTrimStartEmail = strEmail.TrimStart(',');
            //string strTrimEndEmail = strTrimStartEmail.TrimEnd(',');

            //string[] strEmails = strTrimEndEmail.Split(',');

            //foreach (string Email in strEmails)
            //{
            //    //create to address
            //    mmAutomatedMessage.To.Add(Email);
            //}
            ////create subject

            //needs to get back When wee release
            mmAutomatedMessage.To.Add("krishna@sentel.co.uk");
            mmAutomatedMessage.Subject = "Pro Error Log";

            //create the plain text version of the email

            string strBodyText = @"Dear Service Deliver Team,"
                + " " + "Schedule ID" + id1.ToString() + " - " + Reportname + " " + "was not delivered to" + strEmail + "due to blank data being returned. Please check automated reports table.";

            //            string strBodyText = @"Dear Service Deliver Team, 
            //
            //            Please find enlosed an error log of the blank reports:
            //
            //            Many Thanks,
            //            Pro Team 
            //            Sentel Indepedant Ltd 
            //            15 Mckibbin House 
            //            Eastbank Road 
            //            Carryduff 
            //            Belfast 
            //            BT8 8BD 
            //
            //            T: +44 (0)28 9081 5555 
            //            F: +44 (0)28 9081 1055 
            //            E: servicedesk@sentel.co.uk  
            //
            //            Keep up to date on http://www.sentel.co.uk 
            //
            //            Follow us on twitter at http://twitter.com/Sentel_Ind 
            //
            //            Follow us on facebook at http://www.facebook.com/pages/Belfast/Sentel-Independent/249597243565";

            //create the media type for the plain text
            string strMediaType = "text/plain";

            //create an alternative view
            AlternateView avPlainText = AlternateView.CreateAlternateViewFromString(strBodyText, null, strMediaType);

            strBodyText = @"<font face ='verdana' size='3'><p>Dear <b><i>Dear Service Deliver Team</i></b>,</p>"
               + "" +
"Schedule ID" + id1.ToString() + " - " + Reportname + " " + "was not delivered to" + strEmail + "due to blank data being returned. Please check automated reports table.";
            //create the html version of the email
            //            strBodyText = @"<font face ='verdana' size='3'><p>Dear <b><i>Dear Service Deliver Team</i></b>,</p>
            //
            //            <p> </p>
            //
            //            <p>Many Thanks,</p>
            //            <p><b>Pro Team</b></p>
            //            <p>Sentel Independent Ltd
            //            <br />15 McKibbin House
            //            <br />Eastbank Road
            //            <br />Carryduff
            //            <br />Belfast
            //            <br />BT8 8BD
            //            </p>
            //            <p><font face = 'verdana' size ='2'>T: +44 (0)28 9081 5555
            //            <br />F: +44 (0)28 9081 1055
            //            <br />E: servicedesk@sentel.co.uk 
            //            </font>
            //            </p></font>
            //            <font face ='verdana' size='1'>
            //            <p>
            //            <a href='http://www.sentel.co.uk'>www.Sentel.co.uk</a>
            //            <br /><img src='http://www.sentelcallmanagerpro.com//images//SentelLogo.png' alt='Sentel' title='Sentel' />
            //            </p>
            //            <p>
            //            Follow Sentel on Twitter
            //            <br /><img src='http://www.sentelcallmanagerpro.com//images//twitter.png' alt='Twitter' title='Twitter' />
            //            <br />Ask us a question or keep up to date with our latest business news and events using Twitter
            //            <br /><a href='http://twitter.com/Sentel_Ind'>Sentel on Twitter</a>
            //            </p>
            //            <p>
            //            Follow Sentel on Facebook
            //            <br /><img src='http://www.sentelcallmanagerpro.com//images//facebook.png' alt='Facebook' title='FaceBook' />
            //            <br />Add us to get all the latest upgrade news and developments on our products using Facebook
            //            <br /><a href='http://www.facebook.com/pages/Belfast/Sentel-Independent/249597243565'>Sentel on Facebook</a>
            //            </p>
            //            </font>";

            //create the media type for the html
            strMediaType = "text/html";

            //create an alternative view
            AlternateView avHTML = AlternateView.CreateAlternateViewFromString(strBodyText, null, strMediaType);

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
                //ErrorText.Append("There has been a problem connecting to the mail sever. \r\n");
                //tbError.Text = ErrorText.ToString();
            }
            //ProgressText.Append("Error log sent to Service Delivery completed \r\n");
            // tbProgress.Text = ProgressText.ToString();


            //File.Delete("c:\\temp\\" + "TemautomatedReportslogfile" + id1 + Reportname + ".txt");
        }

        public void CreateErrorMessage(string strEmail, int id1, DateTime Datecreated, string Reportname, string selectedname)
        {

            //Attachment attach = new Attachment("c:\\temp\\" + "TemautomatedReportslogfile"+ id1 + Reportname + ".txt");

            //create the email message
            MailMessage mmAutomatedMessage = new MailMessage();

            //create a reply to address
            mmAutomatedMessage.ReplyTo = new MailAddress("servicedesk@sentel.co.uk ");

            //set the priority of the mail message to high to make sure it goes out at quickest possible speed
            mmAutomatedMessage.Priority = MailPriority.High;

            //put a read receipt on each report. Can monitor who is looking at their reports
            mmAutomatedMessage.Headers.Add("Disposition-Notification-To", "ProAutomatedReports@sentel.co.uk");

            //create From address
            //mmAutomatedMessage.From = new MailAddress("TEMAutomatedReports@sentel.co.uk");

            //create to address
            //mmAutomatedMessage.To.Add("Glen.adamson@sentel.co.uk");
            mmAutomatedMessage.From = new MailAddress("ProAutomatedReports@sentel.co.uk");

            string sendEmailsFrom = "ProAutomatedReports@sentel.co.uk";
            string sendEmailsFromPassword = "PR1pr2pr3";
            NetworkCredential cred = new NetworkCredential(sendEmailsFrom, sendEmailsFromPassword);

            //string strTrimStartEmail = strEmail.TrimStart(',');
            //string strTrimEndEmail = strTrimStartEmail.TrimEnd(',');

            //string[] strEmails = strTrimEndEmail.Split(',');

            //foreach (string Email in strEmails)
            //{
            //    //create to address
            //    mmAutomatedMessage.To.Add(Email);
            //}
            ////create subject

            //needs to get back When wee release
            mmAutomatedMessage.To.Add("krishna@sentel.co.uk");
            mmAutomatedMessage.Subject = "Pro Error Log";

            //create the plain text version of the email

            string strBodyText = @"Dear Service Deliver Team,"
                + " " + "Schedule ID" + id1.ToString() + " - " + Reportname + " " + "was not delivered to" + strEmail + "Due to error while Binding the data. Please check automated reports table.";

            //            string strBodyText = @"Dear Service Deliver Team, 
            //
            //            Please find enlosed an error log of the blank reports:
            //
            //            Many Thanks,
            //            Pro Team 
            //            Sentel Indepedant Ltd 
            //            15 Mckibbin House 
            //            Eastbank Road 
            //            Carryduff 
            //            Belfast 
            //            BT8 8BD 
            //
            //            T: +44 (0)28 9081 5555 
            //            F: +44 (0)28 9081 1055 
            //            E: servicedesk@sentel.co.uk  
            //
            //            Keep up to date on http://www.sentel.co.uk 
            //
            //            Follow us on twitter at http://twitter.com/Sentel_Ind 
            //
            //            Follow us on facebook at http://www.facebook.com/pages/Belfast/Sentel-Independent/249597243565";

            //create the media type for the plain text
            string strMediaType = "text/plain";

            //create an alternative view
            AlternateView avPlainText = AlternateView.CreateAlternateViewFromString(strBodyText, null, strMediaType);

            strBodyText = @"<font face ='verdana' size='3'><p>Dear <b><i>Dear Service Deliver Team</i></b>,</p>"
               + "" +
"Schedule ID" + id1.ToString() + " - " + Reportname + " " + "was not delivered to" + strEmail + "Due to error while Binding the data. Please check automated reports table.";
            //create the html version of the email
            //            strBodyText = @"<font face ='verdana' size='3'><p>Dear <b><i>Dear Service Deliver Team</i></b>,</p>
            //
            //            <p> </p>
            //
            //            <p>Many Thanks,</p>
            //            <p><b>Pro Team</b></p>
            //            <p>Sentel Independent Ltd
            //            <br />15 McKibbin House
            //            <br />Eastbank Road
            //            <br />Carryduff
            //            <br />Belfast
            //            <br />BT8 8BD
            //            </p>
            //            <p><font face = 'verdana' size ='2'>T: +44 (0)28 9081 5555
            //            <br />F: +44 (0)28 9081 1055
            //            <br />E: servicedesk@sentel.co.uk 
            //            </font>
            //            </p></font>
            //            <font face ='verdana' size='1'>
            //            <p>
            //            <a href='http://www.sentel.co.uk'>www.Sentel.co.uk</a>
            //            <br /><img src='http://www.sentelcallmanagerpro.com//images//SentelLogo.png' alt='Sentel' title='Sentel' />
            //            </p>
            //            <p>
            //            Follow Sentel on Twitter
            //            <br /><img src='http://www.sentelcallmanagerpro.com//images//twitter.png' alt='Twitter' title='Twitter' />
            //            <br />Ask us a question or keep up to date with our latest business news and events using Twitter
            //            <br /><a href='http://twitter.com/Sentel_Ind'>Sentel on Twitter</a>
            //            </p>
            //            <p>
            //            Follow Sentel on Facebook
            //            <br /><img src='http://www.sentelcallmanagerpro.com//images//facebook.png' alt='Facebook' title='FaceBook' />
            //            <br />Add us to get all the latest upgrade news and developments on our products using Facebook
            //            <br /><a href='http://www.facebook.com/pages/Belfast/Sentel-Independent/249597243565'>Sentel on Facebook</a>
            //            </p>
            //            </font>";

            //create the media type for the html
            strMediaType = "text/html";

            //create an alternative view
            AlternateView avHTML = AlternateView.CreateAlternateViewFromString(strBodyText, null, strMediaType);

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
                // ErrorText.Append("There has been a problem connecting to the mail sever. \r\n");
                // tbError.Text = ErrorText.ToString();
            }
            //ProgressText.Append("Error log sent to Service Delivery completed \r\n");
            //tbProgress.Text = ProgressText.ToString();


            //File.Delete("c:\\temp\\" + "TemautomatedReportslogfile" + id1 + Reportname + ".txt");
        }


        public string ColumnValue(object ColumnValue, int ColumnLength, string ColumnType)
        {
            StringBuilder strResult = new StringBuilder();
            strResult.Append(ColumnValue.ToString());

            switch (ColumnType)
            {
                case "datetime":
                    break;
                case "int":
                    break;
                default:

                    for (int i = strResult.Length; i <= ColumnLength; i++)
                    {
                        strResult.Append(" ");
                    }

                    break;
            }
            return strResult.ToString();
        }

        public void BlankReportLog(int iID, string strEmailAddress, DateTime createddate, string strReportName)
        {
            TextWriter twWriteLog = new StreamWriter("c:\\temp\\" + "TemautomatedReportslogfile" + iID + strReportName + ".txt");

            twWriteLog.Write("Schedule ID" + iID.ToString() + " - " + strReportName + " " + "was not delivered to" + strEmailAddress + "due to blank data being returned. Please check automated reports table.");


            twWriteLog.Flush();
            twWriteLog.Close();
            twWriteLog.Dispose();


            // Tryed with the delete method of file but actually not working any more 

            // CreateErrorMessage(strEmailAddress, iID, createddate, strReportName);
            // File.SetAttributes(("c:\\temp\\" + "TemautomatedReportslogfile" + iID + strReportName + ".txt"), FileAttributes.Normal); 

            // File.Delete("c:\\temp\\" + "TemautomatedReportslogfile" + iID + strReportName + ".txt");

        }
        #endregion
        #region "Doc styling and charts.."
        public Doc.Document CreateDocuments(string strReportname, ArrayList alExport, string from, string to)
        {
            Doc.Document dDoc = new Doc.Document();

            dDoc.DefaultPageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;

            dDoc.Info.Title = strReportname;
            dDoc.Info.Subject = "Please see below for your requested report";
            dDoc.Info.Author = "Sentel Independant Limited";

            DefineStyles(dDoc);
            DefineCover(dDoc);
            DefineTableOfContents(dDoc);
            DefineContentSection(dDoc, from, to);
            DefineParagraphs(dDoc);
            DefineTables(dDoc, alExport, strReportname);
            // DefineCharts(dDoc, alExport, strReportname);

            return dDoc;
        }

        public static void DefineStyles(Doc.Document document)
        {
            //Get the predefined style Normal
            Doc.Style style = document.Styles["Normal"];


            style.Font.Name = "Century Gothic";

            style = document.Styles["Heading1"];
            style.Font.Name = "Century Gothic";
            style.Font.Size = 14;
            style.Font.Bold = true;
            style.Font.Color = Doc.Colors.DarkBlue;
            style.ParagraphFormat.PageBreakBefore = true;
            style.ParagraphFormat.SpaceAfter = 6;

            style = document.Styles["Heading2"];
            style.Font.Size = 12;
            style.Font.Bold = true;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 6;

            style = document.Styles["Heading3"];
            style.Font.Size = 10;
            style.Font.Bold = true;
            style.Font.Italic = true;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 3;

            style = document.Styles[Doc.StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", Doc.TabAlignment.Right);

            style = document.Styles[Doc.StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", Doc.TabAlignment.Center);

            //create a new style called TextBox based on style Normal
            style = document.Styles.AddStyle("TextBox", "Normal");
            style.ParagraphFormat.Alignment = Doc.ParagraphAlignment.Justify;
            style.ParagraphFormat.Borders.Width = 2.5;
            style.ParagraphFormat.Borders.Distance = "3pt";
            style.ParagraphFormat.Shading.Color = Doc.Colors.SkyBlue;

            //create a new style called TOC based on style Normal
            style = document.Styles.AddStyle("TOC", "Normal");
            style.ParagraphFormat.AddTabStop("16cm", Doc.TabAlignment.Right, Doc.TabLeader.Dots);
            style.ParagraphFormat.Font.Color = Doc.Colors.Blue;

            //Create a new style called TOC based on style Normal
            style = document.Styles.AddStyle("TOC", "Normal");
            style.ParagraphFormat.AddTabStop("16cm", Doc.TabAlignment.Right, Doc.TabLeader.Dots);
            style.ParagraphFormat.Font.Color = Doc.Colors.Blue;
        }

        public static void DefineCover(Doc.Document document)
        {
            Doc.Section section = document.AddSection();

            //Doc.Paragraph paragraph = section.AddParagraph();
            Doc.Paragraph paragraph = new Doc.Paragraph();
            //paragraph.Format.SpaceAfter = "3cm";



            Shapes.Image image1 = section.AddImage("c:\\temp\\Callmanager Pro Logo.jpg ");

            image1.Width = "22cm";
            image1.Height = "4cm";

            paragraph = section.AddParagraph("Powered by Sentel Independant Limited");
            paragraph.Format.Font.Size = 20;
            paragraph.Format.Font.Color = Doc.Colors.DarkRed;
            paragraph.Format.SpaceBefore = "6cm";
            paragraph.Format.SpaceAfter = "2cm";

            paragraph = section.AddParagraph("Generated date: ");

            paragraph.AddDateField();
            Shapes.Image image = section.AddImage("c:\\temp\\SentelLogo.png");
        }

        public static void DefineTableOfContents(Doc.Document document)
        {
            Doc.Section section = document.LastSection;

            section.AddPageBreak();

            Doc.Paragraph paragraph = section.AddParagraph("Table of Contents");
            paragraph.Format.Font.Size = 14;
            paragraph.Format.Font.Bold = true;
            paragraph.Format.SpaceAfter = 24;
            paragraph.Format.OutlineLevel = Doc.OutlineLevel.Level1;

            paragraph = section.AddParagraph();
            paragraph.Style = "TOC";
            Doc.Hyperlink hyperlink = paragraph.AddHyperlink("Background");
            hyperlink.AddText("Background \t");
            hyperlink.AddPageRefField("Background");

            paragraph = section.AddParagraph();
            paragraph.Style = "TOC";
            hyperlink = paragraph.AddHyperlink("Data");
            hyperlink.AddText("Data \t");
            hyperlink.AddPageRefField("Data");

            // paragraph = section.AddParagraph();
            //paragraph.Style = "TOC";
            //hyperlink = paragraph.AddHyperlink("Charts");
            //hyperlink.AddText("Charts \t");
            //hyperlink.AddPageRefField("Charts");
        }

        public static void DefineContentSection(Doc.Document document, string from, string to)
        {
            Doc.Section section = document.AddSection();
            section.PageSetup.OddAndEvenPagesHeaderFooter = true;
            section.PageSetup.StartingNumber = 1;

            Doc.HeaderFooter header = section.Headers.Primary;
            header.AddParagraph("\t Powered by Sentel");

            header = section.Headers.EvenPage;
            header.AddParagraph("Powered by Sentel");
            header.AddParagraph("Date Range:" + " " + from + "---" + to);

            //temp - can prob be deleted
            //create a paragraph with centered page number.
            Doc.Paragraph paragraph = new Doc.Paragraph();
            paragraph.AddTab();
            paragraph.AddPageField();

            //Add paragraph to footer for odd pages
            section.Footers.Primary.Add(paragraph);

            //add a clone to the even pages.
            section.Footers.EvenPage.Add(paragraph.Clone());
        }

        public static void DefineParagraphs(Doc.Document document)
        {
            Doc.Paragraph paragraph = document.LastSection.AddParagraph("Pro Automated Reports", "Heading1");
            paragraph.AddBookmark("Background");

            SetDisclaimer(document);
        }

        public static void DefineTables(Doc.Document document, ArrayList alExport, string strReportName)
        {
            MigraDoc.DocumentObjectModel.Paragraph paragraph = document.LastSection.AddParagraph(strReportName, "Heading1");
            paragraph.AddBookmark("Data");

            CreateTable(document, alExport);
        }

        public static void CreateTable(Doc.Document document, ArrayList alExport)
        {
            document.LastSection.AddParagraph("Report", "Heading2");

            Table table = new Table();
            table.Borders.Width = 1;
            table.Rows.Height = 2;
            table.Columns.Width = 100;

            Column column = new Column();

            int iCounter = 0;

            do
            {
                table.AddColumn();
                //(Unit.FromCentimeter(2.5));
                //(Unit.FromCentimeter(5));
                iCounter++;
            }

            while (alExport[iCounter].ToString() != "\r\n");

            iCounter = 0;



            //table.AddColumn(Unit.FromCentimeter(2));
            //column.Format.Alignment = ParagraphAlignment.Center;

            //table.AddColumn(Unit.FromCentimeter(5));
            //table.AddColumn(Unit.FromCentimeter(5));

            Row row = table.AddRow();
            row.Shading.Color = Colors.LightBlue;
            int iRowCount = 1;
            int iColumnCount = 0;

            for (int i = 0; i < alExport.Count; i++)
            {
                Cell cell = new Cell();

                if (alExport[i].ToString() == "\r\n")
                {
                    if (i < (alExport.Count - 1))
                    {
                        iColumnCount = iCounter;
                        row = table.AddRow();
                        iRowCount += 1;
                    }

                    iCounter = 0;
                }
                else
                {
                    cell = row.Cells[iCounter];

                    cell.AddParagraph(alExport[i].ToString());
                    iCounter++;
                }
            }

            table.SetEdge(0, 0, iColumnCount, iRowCount, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 3, Colors.PaleVioletRed);
            //                (0, 0, 3, 3, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 1.5, Colors.Black);

            document.LastSection.Add(table);
        }

        public static void DefineCharts(Doc.Document document, ArrayList alExport, string strReportName)
        {
            MigraDoc.DocumentObjectModel.Paragraph paragraph = document.LastSection.AddParagraph(strReportName + "Chart", "Heading1");
            paragraph.AddBookmark("Charts");

            document.LastSection.AddParagraph(strReportName + "Chart", "Heading2");

            Doc.Shapes.Charts.Chart chart = new Doc.Shapes.Charts.Chart();
            chart.Left = 0;

            chart.Width = Unit.FromCentimeter(25);
            chart.Height = Unit.FromCentimeter(12);
            Doc.Shapes.Charts.Series series = chart.SeriesCollection.AddSeries();

            int xaxis = 10;
            int yaxis = 10;
            int iColumnCount = 0;
            bool blnExitForLoop = false;

            XSeries xseries = chart.XValues.AddXSeries();
            series.ChartType = ChartType.Column2D;
            series.HasDataLabel = true;
            List<Double> lTime = new List<double>();
            int iLoopCtr = 0;

            for (int i = 0; i <= alExport.Count; i++)
            {
                if (alExport[i].ToString() == "\r\n")
                {
                    i = alExport.Count + 1;
                }
                iLoopCtr++;
            }

            for (int i = iLoopCtr; i <= alExport.Count; i++)
            {
                if (alExport[i].ToString() == "\r\n")
                {
                    iColumnCount = 0;
                    if (i < alExport.Count)
                    {
                        i++;
                    }
                    else
                    {
                        i = alExport.Count + 1;
                        blnExitForLoop = true;
                    }
                }

                if (blnExitForLoop == false)
                {
                    if (iColumnCount == xaxis)
                    {
                        xseries.Add(alExport[i].ToString());
                    }
                    else if (iColumnCount == yaxis)
                    {

                        lTime.Add(double.Parse(alExport[i].ToString()));
                    }

                    iColumnCount++;
                }
            }

            series.Add(lTime.ToArray());

            chart.XAxis.MajorTickMark = TickMarkType.Outside;
            chart.XAxis.Title.Caption = "Employee";

            chart.YAxis.MajorTickMark = TickMarkType.Outside;
            chart.YAxis.HasMajorGridlines = true;

            chart.PlotArea.LineFormat.Color = Colors.DarkGray;
            chart.PlotArea.LineFormat.Width = 1;

            document.LastSection.Add(chart);
        }

        public void CreateBarChart(SqlCommand oCmdCDR, SqlDataReader drReader, string strDateTime, string strXaxis, string strYaxis, string ReportName)
        {
            drReader = oCmdCDR.ExecuteReader();
            drReader.Read();

            Dundas.Charting.WinControl.Chart chart1 = new Dundas.Charting.WinControl.Chart();
            chart1.Titles.Add("Title1");
            chart1.Titles.Add("Title2");
            chart1.Titles["Title1"].Text = "PRO";
            chart1.Titles["Title1"].Color = System.Drawing.Color.FromArgb(26, 59, 105);
            chart1.Titles["Title1"].Position.Height = 6;
            chart1.Titles["Title1"].Position.Width = 89.30075F;
            chart1.Titles["Title1"].Position.X = 4.879699F;
            chart1.Titles["Title1"].Position.Y = 4;

            chart1.Titles["Title2"].Text = ReportName;
            chart1.Titles["Title2"].Color = System.Drawing.Color.FromArgb(26, 59, 105);
            chart1.Titles["Title2"].Position.Height = 6;
            chart1.Titles["Title2"].Position.Width = 89.30075F;
            chart1.Titles["Title2"].Position.X = 4.879699F;
            chart1.Titles["Title2"].Position.Y = 12;
            chart1.Legends.Add("Default");
            chart1.Legends["Default"].AutoFitText = false;

            chart1.Series.Add(strXaxis);
            chart1.Series[strXaxis].BorderColor = System.Drawing.Color.FromArgb(180, 26, 59, 105);
            chart1.Series[strXaxis].MarkerBorderWidth = 2;
            chart1.Series[strXaxis].ShowLabelAsValue = true;
            chart1.Series[strXaxis].XValueType = ChartValueTypes.Double;
            chart1.Series[strXaxis].YValueType = ChartValueTypes.Double;

            chart1.ChartAreas.Add("Default1");
            chart1.ChartAreas["Default"].AxisX.Title = strXaxis;
            chart1.ChartAreas["Default"].AxisY.Title = strYaxis;
            chart1.ChartAreas["Default"].BackGradientEndColor = System.Drawing.Color.Transparent;
            chart1.ChartAreas["Default"].BackGradientType = GradientType.TopBottom;
            chart1.ChartAreas["Default"].BorderColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart1.ChartAreas["Default"].ShadowColor = System.Drawing.Color.Transparent;
            chart1.ChartAreas["Default"].AxisY.LabelsAutoFit = false;
            chart1.ChartAreas["Default"].AxisY.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart1.ChartAreas["Default"].AxisY.MajorGrid.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart1.ChartAreas["Default"].AxisY.MajorTickMark.Enabled = false;
            chart1.ChartAreas["Default"].AxisY.LabelStyle.Enabled = false;
            chart1.ChartAreas["Default"].AxisX.LabelsAutoFit = false;
            chart1.ChartAreas["Default"].AxisX.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart1.ChartAreas["Default"].AxisX.MajorGrid.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart1.ChartAreas["Default"].AxisX.MajorTickMark.Enabled = false;
            chart1.ChartAreas["Default"].AxisX.LabelStyle.Enabled = false;
            chart1.ChartAreas["Default"].AxisX.Interval = 1;
            chart1.ChartAreas["Default"].AxisY.LabelsAutoFit = false;
            chart1.ChartAreas["Default"].AxisY.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart1.ChartAreas["Default"].AxisY.MajorGrid.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart1.ChartAreas["Default"].AxisY.MajorTickMark.Enabled = false;
            chart1.ChartAreas["Default"].Position.Height = 70;
            chart1.ChartAreas["Default"].Position.Width = 94;
            chart1.ChartAreas["Default"].Position.X = 3;
            chart1.ChartAreas["Default"].Position.Y = 24;
            chart1.ChartAreas["Default"].Area3DStyle.Clustered = true;
            chart1.ChartAreas["Default"].Area3DStyle.Perspective = 10;
            chart1.ChartAreas["Default"].Area3DStyle.RightAngleAxes = false;
            chart1.ChartAreas["Default"].Area3DStyle.WallWidth = 0;
            chart1.ChartAreas["Default"].Area3DStyle.XAngle = 15;
            chart1.ChartAreas["Default"].Area3DStyle.YAngle = 10;
            chart1.Visible = true;
            chart1.Height = 348;
            chart1.Width = 460;
            chart1.Palette = ChartColorPalette.Dundas;
            chart1.BackColor = System.Drawing.Color.FromArgb(221, 223, 240);
            chart1.ChartAreas[0].BackColor = System.Drawing.Color.FromArgb(255, 255, 255, 255);
            chart1.BackGradientType = GradientType.TopBottom;
            chart1.BorderLineColor = System.Drawing.Color.FromArgb(26, 59, 105);
            chart1.BorderLineStyle = ChartDashStyle.Solid;
            chart1.BorderLineWidth = 2;

            chart1.Series[strXaxis].Points.DataBindXY(drReader, strXaxis, drReader, strYaxis);

            chart1.SaveAsImage("c:\\visualstudio\\projects\\TemAutomatedReports\\images\\chart" + strDateTime + ".png", ChartImageFormat.Png);

            chart1.Series.Remove(strXaxis);
            chart1.Series.Remove(strYaxis);
            chart1.Dispose();

            drReader.Close();
        }

        public void CreatePieChart(SqlCommand oCmdCDR, SqlDataReader drReader, string strDateTime, string strXaxis, string strYaxis, string ReportName)
        {
            drReader = oCmdCDR.ExecuteReader();
            drReader.Read();

            Dundas.Charting.WinControl.Chart chart2 = new Dundas.Charting.WinControl.Chart();

            chart2.Titles.Add("Title1");
            chart2.Titles["Title1"].Text = "PRO";
            chart2.Titles["Title1"].Color = System.Drawing.Color.FromArgb(26, 59, 105);
            chart2.Titles["Title1"].Position.Height = 6;
            chart2.Titles["Title1"].Position.Width = 89.30075F;
            chart2.Titles["Title1"].Position.X = 4.879699F;
            chart2.Titles["Title1"].Position.Y = 4;

            chart2.Titles.Add("Title2");
            chart2.Titles["Title2"].Text = ReportName;
            chart2.Titles["Title2"].Color = System.Drawing.Color.FromArgb(26, 59, 105);
            chart2.Titles["Title2"].Position.Height = 5;
            chart2.Titles["Title2"].Position.Width = 89.30075F;
            chart2.Titles["Title2"].Position.X = 4.879699F;
            chart2.Titles["Title2"].Position.Y = 12;

            chart2.Legends.Add("Default");
            chart2.Legends["Default"].AutoFitText = false;
            chart2.Legends["Default"].BackColor = System.Drawing.Color.Transparent;
            chart2.Legends["Default"].Docking = LegendDocking.Left;
            chart2.Legends["Default"].Position.Height = 38;
            chart2.Legends["Default"].Position.Width = 30;
            chart2.Legends["Default"].Position.X = 3;
            chart2.Legends["Default"].Position.Y = 29;

            chart2.Series.Add(strXaxis);
            chart2.Series[strXaxis].BorderColor = System.Drawing.Color.FromArgb(50, 0, 0, 0);

            chart2.Series[strXaxis].MarkerBorderWidth = 2;
            chart2.Series[strXaxis].ShowLabelAsValue = true;
            chart2.Series[strXaxis].XValueType = ChartValueTypes.Double;
            chart2.Series[strXaxis].YValueType = ChartValueTypes.Double;
            chart2.Series[strXaxis].Type = SeriesChartType.Pie;

            chart2.ChartAreas.Add("Default");
            chart2.ChartAreas["Default"].AxisX.Title = strXaxis;
            chart2.ChartAreas["Default"].AxisY.Title = strYaxis;
            chart2.ChartAreas["Default"].BackGradientEndColor = System.Drawing.Color.Transparent;
            chart2.ChartAreas["Default"].BackGradientType = GradientType.TopBottom;
            chart2.ChartAreas["Default"].BackColor = System.Drawing.Color.Transparent;
            chart2.ChartAreas["Default"].BorderColor = System.Drawing.Color.Transparent;
            chart2.ChartAreas["Default"].ShadowColor = System.Drawing.Color.Transparent;
            chart2.ChartAreas["Default"].AxisY.LabelsAutoFit = false;
            chart2.ChartAreas["Default"].AxisY.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart2.ChartAreas["Default"].AxisY.MajorGrid.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart2.ChartAreas["Default"].AxisY.MajorTickMark.Enabled = false;
            chart2.ChartAreas["Default"].AxisY.LabelStyle.Enabled = false;
            chart2.ChartAreas["Default"].AxisX.LabelsAutoFit = false;
            chart2.ChartAreas["Default"].AxisX.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart2.ChartAreas["Default"].AxisX.MajorGrid.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart2.ChartAreas["Default"].AxisX.MajorTickMark.Enabled = false;
            chart2.ChartAreas["Default"].AxisX.LabelStyle.Enabled = false;
            chart2.ChartAreas["Default"].AxisX.Interval = 1;
            chart2.ChartAreas["Default"].AxisY.LabelsAutoFit = false;
            chart2.ChartAreas["Default"].AxisY.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart2.ChartAreas["Default"].AxisY.MajorGrid.LineColor = System.Drawing.Color.FromArgb(64, 64, 64, 64);
            chart2.ChartAreas["Default"].AxisY.MajorTickMark.Enabled = false;
            chart2.ChartAreas["Default"].Position.Height = 70;
            chart2.ChartAreas["Default"].Position.Width = 94;
            chart2.ChartAreas["Default"].Position.X = 3;
            chart2.ChartAreas["Default"].Position.Y = 24;
            chart2.ChartAreas["Default"].Area3DStyle.Clustered = true;
            chart2.ChartAreas["Default"].Area3DStyle.Perspective = 10;
            chart2.ChartAreas["Default"].Area3DStyle.RightAngleAxes = false;
            chart2.ChartAreas["Default"].Area3DStyle.WallWidth = 0;
            chart2.ChartAreas["Default"].Area3DStyle.XAngle = 15;
            chart2.ChartAreas["Default"].Area3DStyle.YAngle = 10;

            chart2.Visible = true;
            chart2.Height = 348;
            chart2.Width = 460;
            chart2.Palette = ChartColorPalette.Dundas;
            chart2.BackColor = System.Drawing.Color.FromArgb(221, 223, 240);
            chart2.ChartAreas[0].BackColor = System.Drawing.Color.FromArgb(255, 255, 255, 255);
            chart2.BackGradientType = GradientType.TopBottom;
            chart2.BorderLineColor = System.Drawing.Color.FromArgb(26, 59, 105);
            chart2.BorderLineStyle = ChartDashStyle.Solid;
            chart2.BorderLineWidth = 2;

            chart2.Series[strXaxis].Points.DataBindXY(drReader, strXaxis, drReader, strYaxis);

            chart2.PaletteCustomColors = new System.Drawing.Color[]
           {
               System.Drawing.Color.FromArgb(185, 11, 0),     //TEMReport
               System.Drawing.Color.FromArgb(201, 49, 0),     //Dashboard
               System.Drawing.Color.FromArgb(219, 65, 3),     //Reporting
               System.Drawing.Color.FromArgb(228, 205, 27),   //KPI
               System.Drawing.Color.FromArgb(198, 217, 1),    //Billing
               System.Drawing.Color.FromArgb(49, 195, 0),     //Voice Recording
               System.Drawing.Color.FromArgb(0, 193, 190),    //Trend
               System.Drawing.Color.FromArgb(0, 158, 213),    //Directory
               System.Drawing.Color.FromArgb(0, 82, 216),     //Inventory
               System.Drawing.Color.FromArgb(152,4,214),      //Tariff
               System.Drawing.Color.FromArgb(73, 4, 181)      //Admin
           };

            chart2.SaveAsImage("c:\\visualstudio\\projects\\TemAutomatedReports\\images\\pie" + strDateTime + ".png", ChartImageFormat.Png);

            chart2.Series.Remove(strXaxis);
            chart2.Series.Remove(strYaxis);
            chart2.Dispose();
            drReader.Close();

        }

        private string GetDataTableAsHTML(DataTable thisTable, string strColumns, ArrayList unwantedlist)
        {
            ArrayList alExport = new ArrayList();
            string[] Columns = strColumns.Split(',');

            //15,5,8,79,46,
            //string [] Length = strColumns.Split('@')


            foreach (string Column in Columns)
            {
                string[] arrColumnNameAndLength = Column.Split('@', '#');

                //do function and pass in column name and amount
                alExport.Add(arrColumnNameAndLength[0]);

            }
            //foreach (string un in unwantedlist)
            //{

            //    alExport.Remove(un);

            //}

            // the code need to change 

            //foreach (DataColumn dc in thisTable.Columns)
            //{
            //    if (dc.ColumnName != alExport.ToString())
            //    {
            //        thisTable.Columns.Remove(dc.ColumnName);
            //    }
            //}

            System.Text.StringBuilder sb = new System.Text.StringBuilder();

            sb.Append("<TR bgcolor= '#CAD8EC'>");

            //first append the column names.
            foreach (DataColumn column in thisTable.Columns)
            {

                foreach (string somecolumn in alExport)
                {
                    if (column.ColumnName != somecolumn)
                    {

                    }
                    else
                    {
                        sb.Append("<TD><B>");
                        sb.Append(column.ColumnName);
                        sb.Append("</B></TD>");
                    }
                }
            }

            sb.Append("</TR>");

            // next, the column values.
            foreach (DataRow row in thisTable.Rows)
            {
                sb.Append("<TR>");

                foreach (DataColumn column in thisTable.Columns)
                {
                    foreach (string somecolumn in alExport)
                    {
                        if (column.ColumnName != somecolumn)
                        {

                        }
                        else
                        {
                            sb.Append("<TD>");
                            if (row[column].ToString().Trim().Length > 0)
                                sb.Append(row[column]);
                            else
                                sb.Append(" ");
                            sb.Append("</TD>");
                        }
                    }
                }

                sb.Append("</TR>");
            }
            return sb.ToString();
        }

        public String GenerateHTMLCDRResults(string strStoredProcedure, string strListofFilters, string strEmailAddresses, string strColumns, string Chosennodelist, int id1, ArrayList pk)
        {
            string strData = "";


            // This data is commented because it is based on hourly schedule report 

            //strFromDate = DateTime.Now.Year.ToString() + "-" +
            //    PadIntWithZeros(DateTime.Now.Month.ToString()) + "-" +
            //    PadIntWithZeros(DateTime.Now.Day.ToString()) + " " +
            //    PadIntWithZeros(DateTime.Now.AddHours(-1).Hour.ToString()) + ":00:00";
            ////"13:00:00";

            //strToDate = DateTime.Now.Year.ToString() + "-" +
            //    PadIntWithZeros(DateTime.Now.Month.ToString()) + "-" +
            //    PadIntWithZeros(DateTime.Now.Day.ToString()) + " " +
            //   PadIntWithZeros(DateTime.Now.Hour.ToString()) + ":00:00";
            ////"14:00:00";

            //strListofFilters = strListofFilters.Replace("@FromDate=","@FromDate='" + strFromDate + "', @ToDate='" + strToDate + "'");

            //This List Should be read from the Database

            DataSet dsCDRResults = new DataSet();
            using (SqlConnection oConn = new SqlConnection(TEMConnectionString))
            {


                using (SqlCommand oCmd = new SqlCommand("exec " + strStoredProcedure + " " + strListofFilters + "," + Chosennodelist, oConn))
                {
                    SqlDataAdapter sqlDA = new SqlDataAdapter(oCmd);
                    oCmd.CommandTimeout = 3000;

                    sqlDA.Fill(dsCDRResults);
                }

                oConn.Close();
            }

            //TODO; pass dsCDRResults.Table[1] aswell rather than pk
            strData = GetDataTableAsHTML(dsCDRResults.Tables[0], strColumns, pk);

            return strData;

        }

        static void SetDisclaimer(Doc.Document document)
        {
            Doc.Paragraph paragraph = document.LastSection.AddParagraph();
            paragraph.Format.Alignment = Doc.ParagraphAlignment.Left;
            paragraph.AddText(@"We are so confident in our Pro approach that we offer a 
            service where we guarantee net savings for our customers. Our wide range of 
            Pro services, include a tariff analyzer, invoice reconciliation, asset 
            register and analysis, validation, authorisation and charge back of bills, 
            moves, adds, changes and disconnects, sourcing comms packages and comparing 
            against databases of industry best-in-class tariffs and charges, help desk, 
            and other services.");
        }
        #endregion
        #region "Port folio report"
        private void InsertPortfoliodetails(string filename, string strEmailAddress)
        {

            string[] email = strEmailAddress.Split(',');

            foreach (string s in email)
            {
                if (s.Contains('@'))
                {
                    var pfreports = new tbl_PortfolioReport();
                    pfreports.PortfolioReport_ReportName = filename;
                    pfreports.PortfolioReport_Email = s;
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

        private List<int> Getportfolioreports()
        {

            List<int> Reports = new List<int>();
            var rep = DbContext.tbl_customportfolioreports.Select(s => s.schedule_id_PK);

            foreach (var v in rep)
            {

                Reports.Add(Convert.ToInt32(v));

            }
            return Reports;

        }

        private void updatecustomportfolio(string filename, int id)
        {
            var v = DbContext.tbl_customportfolioreports.Where(r => r.schedule_id_PK == id).First();
            v.Customportfolio_Report = filename;
            v.Customportfolio_Run = true;
            try
            {
                DbContext.SubmitChanges();
            }

            catch
            { }

        }
        #endregion      
        # region "Costcenter methods"
        private DataSet ExecuteRawData(Schedule schreports)
        {
            string[] choosennode = schreports.Chosennodelist.Split('=');
            string param = "'" + choosennode[1] + "'";
            string name = choosennode[0] + "=";

            schreports.Chosennodelist = name + param;
            SqlConnection oConn = new SqlConnection(TEMConnectionString);
            SqlCommand oCmd = new SqlCommand("exec " + schreports.StoredProcedureName + " " + schreports.ListofFilters + "," + schreports.Chosennodelist, oConn);
            // need to change the choosennodelist with ''
            DataSet tables = new DataSet();
            try
            {
                SqlDataAdapter sqlDA = new SqlDataAdapter(oCmd);

                oCmd.CommandTimeout = 3000;
                sqlDA.Fill(tables);
            }
            catch
            {

            }
            finally
            { oConn.Close(); }


            return tables;


        }
        public void GenerateCostcenterreport(Schedule schReports)
        {

            DataSet tables = ExecuteRawData(schReports );

            if (tables.Tables[0].Rows.Count > 0)
            {
                string[] tofromdates1 = schReports.ListofFilters.Split('=', ',');
                string trimmedstartfromdate = tofromdates1[5].TrimStart('\'');
                string trimmedfromdate1 = trimmedstartfromdate.TrimEnd('\'');
                string trimmedstarttodate = tofromdates1[7].TrimStart('\'');
                string trimmedtodate1 = trimmedstarttodate.TrimEnd('\'');

                switch (schReports.Type)
                {
                    case "html":
                        PrepareHTMLreport(schReports, tables, trimmedstartfromdate, trimmedstarttodate, "4");
                        break;
                    case "word":
                        PrepareWORDreport(schReports, tables, trimmedstartfromdate, trimmedstarttodate, "4");
                        break;
                    case "text":
                        PrepareTEXTreport(schReports, tables, trimmedstartfromdate, trimmedstarttodate, "4");
                        break;
                    case "PDF":
                        PreparePDFreport(schReports, tables, trimmedstartfromdate, trimmedstarttodate, "3");
                        break;
                    default:
                        PrepareCSVreport(schReports, tables, trimmedstartfromdate, trimmedstarttodate, "3");
                        break;
                }

            }
            else
            { }// should return no data message......   }

        }
        private int GetLevels(DataSet dt)
        {
            if (dt.Tables[6].Rows.Count > 1)
            {
                Level = 4;
                return 4;
            }
            else if (dt.Tables[5].Rows.Count > 1)
            {
                Level = 3;
                return 3;


            }
            else if (dt.Tables[4].Rows.Count > 1)
            {
                Level = 2;
                return 2;

            }
            else
            {
                Level = 1;
                return 1;
            }


        }
        private void PrepareHTMLstart(string trimmedstartfromdate, string trimmedstarttodate, TextWriter twWriter, Schedule schReports)
        {

            twWriter.Write("<HTML>");
            twWriter.Write("<HEAD>");
            twWriter.Write("<TITLE>");
            twWriter.Write(schReports.ReportName);
            twWriter.Write("</TITLE>");
            twWriter.Write("<style type='text/css'>");
            twWriter.Write(".style1");
            twWriter.Write("{");
            twWriter.Write("font-size: 70pt;");
            twWriter.Write("}");
            twWriter.Write("</style>");
            twWriter.Write("</HEAD>");
            twWriter.Write("<BODY BGCOLOR='white'>");
            twWriter.Write("<CENTER>");
            twWriter.Write("<img src='http://goo.gl/3NmNM' />");
            twWriter.Write("<BR><BR><H2>" + schReports.ReportName + "</H2>");
            #region may need removed from code if table does not work
            twWriter.Write("<CENTER>");
            twWriter.Write("<table cellpadding='0' cellspacing='0' border='0' class='opaque'>");
            twWriter.Write("<TR>");
            twWriter.Write("<TD>");
            twWriter.Write("</TD>");
            twWriter.Write("<TD>");
            twWriter.Write("</TD>");
            twWriter.Write("<TD align ='left'>");
            twWriter.Write("Date Range :  ");
            twWriter.Write(trimmedstartfromdate);
            twWriter.Write("  --  ");
            twWriter.Write(trimmedstarttodate);
            twWriter.Write("</TD>");
            twWriter.Write("<TD>");
            twWriter.Write("</TD>");
            twWriter.Write("</TR>");
            twWriter.Write("<CENTER>");
            twWriter.Write("<tr>");
            twWriter.Write("<td style='width: 12px; height: 1px;'><img alt='' src='http://goo.gl/1ERss' width='12' height='1' /></td>");
            twWriter.Write("<td style='width: 18px; height: 16px; background-image: url(http://goo.gl/F1Y6K);'><img alt='' src='http://goo.gl/1ERss' width='18' height='16' /></td>");
            twWriter.Write("<td style='width: 910px; height: 16px; background-image: url(http://goo.gl/twveK);'><img alt='' src='http://goo.gl/1ERss' width='910' height='16' /></td>");
            twWriter.Write("<td style='width: 23px; height: 16px; background-image: url(http://goo.gl/kCStA);'><img alt='' src='http://goo.gl/1ERss' width='23' height='16' /></td>");
            twWriter.Write("</tr>");
            twWriter.Write("<tr>");
            twWriter.Write("<td style='width: 12px; height: 1px;'><img alt='' src='http://goo.gl/1ERss' width='12' height='1' /></td>");
            twWriter.Write("<td style='width: 18px; background-image: url(http://goo.gl/sJxcZ);'></td>");
            twWriter.Write("<td>");

        }
        private void PrepareHTMLend(TextWriter twWriter)
        {

            twWriter.Write("</TD></TR>");
            twWriter.Write("</table >");
            twWriter.Write("</td>");
            twWriter.Write("<td style='width: 23px; background-image: url(http://goo.gl/HQSbi);'></td>");
            twWriter.Write("</tr>");
            twWriter.Write("<tr>");
            twWriter.Write("<td style='width: 12px; height: 1px;'><img alt='' src='http://goo.gl/1ERss' width='12' height='1' /></td>");
            twWriter.Write("<td style='width: 18px; height: 23px; background-image: url(http://goo.gl/aZ7T4);'><img alt='' src='http://goo.gl/1ERss' width='18' height='23' /></td>");
            twWriter.Write("<td style='width: 910px; height: 23px; background-image: url(http://goo.gl/Ljz7E);'><img alt='' src='http://goo.gl/1ERss' width='910' height='23' /></td>");
            twWriter.Write("<td style='width: 23px; height: 23px; background-image: url(http://goo.gl/caJ7u);'><img alt='' src='http://www.sentelcallmanagerpro.com/images/spacer.gif' width='23' height='23' /></td>");
            twWriter.Write("</tr>");
            twWriter.Write("</table>");
            twWriter.Write("</table>");


        }
        private string WriteHTMlTable(DataTable thisTable, int value)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            if (value == 0)
            {
                thisTable.Columns.RemoveAt(0);
            }

            sb.Append("<TR bgcolor= '#CAD8EC'>");

            //first append the column names.
            foreach (DataColumn column in thisTable.Columns)
            {

                sb.Append("<TD><B>");
                sb.Append(column.ColumnName);
                sb.Append("</B></TD>");
            }

            sb.Append("</TR>");

            // next, the column values.
            foreach (DataRow row in thisTable.Rows)
            {
                sb.Append("<TR>");

                foreach (DataColumn column in thisTable.Columns)
                {
                    sb.Append("<TD>");
                    if (row[column].ToString().Trim().Length > 0)
                        sb.Append(row[column]);
                    else
                        sb.Append(" ");
                    sb.Append("</TD>");
                }

                sb.Append("</TR>");
            }
            return sb.ToString();

        }
        private StringBuilder WriteHTMlextTable(DataTable des, DataTable dt)
        {




            var depts = dt.AsEnumerable().Select(s => s.Field<string>("CostCentre")).Distinct();
            StringBuilder HTMLtoRender = new StringBuilder();
            foreach (var v in depts)
            {
                HTMLtoRender.Append("<tr>");
                HTMLtoRender.AppendFormat("<td align = \"center\" style=\"color:red;font-weight:bold;\">");
                HTMLtoRender.Append(v.ToString());
                HTMLtoRender.AppendFormat("</td>");
                HTMLtoRender.Append("</tr>");

                HTMLtoRender.Append("<tr  bgcolor= '#CAD8EC' >");
                foreach (DataColumn column in des.Columns)
                {
                    if (column.ColumnName != "Department")
                    {
                        HTMLtoRender.AppendFormat("<td ><B>" + column.ColumnName + "</B></td>");
                    }

                }
                HTMLtoRender.Append("</tr>");

                var dest = from db in des.AsEnumerable()
                           where db.Field<string>("Department") == v.ToString()
                           select new
                           {
                               Destinationname = db.Field<object>("Destinationname"),
                               Totalcalls = db.Field<object>("Totalcalls"),
                               Totalduration = db.Field<object>("Totalduration"),
                               Cost = db.Field<object>("Cost")

                           };
                foreach (var destinations in dest)
                {
                    HTMLtoRender.Append("</tr>");
                    HTMLtoRender.AppendFormat("<td>", "150px");
                    HTMLtoRender.Append(destinations.Destinationname.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.AppendFormat("<td >", "150px");
                    HTMLtoRender.Append(destinations.Totalcalls.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.AppendFormat("<td>", "150px");
                    HTMLtoRender.Append(destinations.Totalduration.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.AppendFormat("<td>", "150px");
                    HTMLtoRender.Append(destinations.Cost.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");

                    HTMLtoRender.Append("</tr>");
                }
                HTMLtoRender.Append("</BR>");
                HTMLtoRender.Append("<tr  bgcolor= '#CAD8EC' >");
                foreach (DataColumn column in dt.Columns)
                {
                    if (column.ColumnName != "CostCentre")
                    {
                        HTMLtoRender.AppendFormat("<td ><B>" + column.ColumnName + "</B></td>");
                    }


                }
                HTMLtoRender.Append("</tr>");


                var records = from db in dt.AsEnumerable()
                              where db.Field<string>("CostCentre") == v.ToString()
                              select new
                              {
                                  Name = db.Field<object>("Name"),
                                  Extension = db.Field<object>("Extension"),
                                  OutgoingCalls = db.Field<object>("OutgoingCalls"),
                                  OutgoingDuration = db.Field<object>("OutgoingDuration"),
                                  RingResponse = db.Field<object>("RingResponse"),
                                  AbandonedCalls = db.Field<object>("AbandonedCalls"),
                                  IncomingCalls = db.Field<object>("IncomingCalls"),
                                  IncomingDuration = db.Field<object>("IncomingDuration"),
                                  Cost = db.Field<object>("Cost")

                              };
                foreach (var all in records)
                {
                    HTMLtoRender.Append("<tr>");
                    HTMLtoRender.AppendFormat("<td>", "150px");
                    HTMLtoRender.Append(all.Name.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.AppendFormat("<td>", "150px");
                    HTMLtoRender.Append(all.Extension.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");

                    HTMLtoRender.AppendFormat("<td >", "150px");
                    HTMLtoRender.Append(all.RingResponse.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");

                    HTMLtoRender.AppendFormat("<td >", "150px");
                    HTMLtoRender.Append(all.IncomingCalls.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.AppendFormat("<td >", "150px");
                    HTMLtoRender.Append(all.IncomingDuration.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.AppendFormat("<td>", "150px");
                    HTMLtoRender.Append(all.AbandonedCalls.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.AppendFormat("<td>", "150px");
                    HTMLtoRender.Append(all.OutgoingCalls.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.AppendFormat("<td >", "150px");
                    HTMLtoRender.Append(all.OutgoingDuration.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.AppendFormat("<td>", "150px");
                    HTMLtoRender.Append(all.Cost.ToString().Replace(" ", string.Empty));
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.Append("</tr>");
                }




            }

            // HTMLtoRender.Append("</table>");
            return HTMLtoRender;
        }
        // This method is to create HTML report version
        private void PrepareHTMLreport(Schedule schReports, DataSet dt, string trimmedstartfromdate, string trimmedstarttodate, string level)
        {
            TextWriter twWriter = new StreamWriter("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + schReports.ID1 + ".html");


            PrepareHTMLstart(trimmedstartfromdate, trimmedstarttodate, twWriter, schReports);
            twWriter.Write("<table WIDTH='100%' cellpadding='5' cellspacing='0'align ='center'>");
            twWriter.Write("<CENTER>");
            twWriter.Write("<TR bgcolor= '#CAD8EC'>");

            twWriter.Write("</TR>");
            twWriter.Write("<TR bgcolor= '#CAD8EC'><TD>");

            // string level = "3";//tofromdates1[93];

            if (level == "4")
            {
                twWriter.Write(WriteHTMlTable(dt.Tables[2], 0));
                int tablevalue = GetLevels(dt);
                if (tablevalue != 1)
                {
                    twWriter.Write(WriteHTMlTable(dt.Tables[GetLevels(dt) + 2], 1));
                }
                else { twWriter.Write(WriteHTMlextTable(dt.Tables[0], dt.Tables[3])); }
            }
            else
            {
                if (GetLevels(dt) == 1)
                {
                    twWriter.Write(WriteHTMlTable(dt.Tables[3], 0));
                }
                else if (GetLevels(dt) == 2)
                {
                    twWriter.Write(WriteHTMlTable(dt.Tables[2], 0));
                    twWriter.Write(WriteHTMlTable(dt.Tables[4], 0));
                    twWriter.Write(WriteHTMlextTable(dt.Tables[0], dt.Tables[3]));
                }
                else if (GetLevels(dt) == 3)
                {
                    twWriter.Write(WriteHTMlTable(dt.Tables[2], 0));
                    twWriter.Write(WriteHTMlTable(dt.Tables[5], 1));
                    twWriter.Write(WriteHTMlTable(dt.Tables[4], 1));
                    twWriter.Write(WriteHTMlextTable(dt.Tables[0], dt.Tables[3]));
                }
                else if (GetLevels(dt) == 4)
                {
                    twWriter.Write(WriteHTMlTable(dt.Tables[2], 0));
                    twWriter.Write(WriteHTMlTable(dt.Tables[6], 1));
                    twWriter.Write(WriteHTMlTable(dt.Tables[5], 1));
                    twWriter.Write(WriteHTMlTable(dt.Tables[4], 1));
                    twWriter.Write(WriteHTMlextTable(dt.Tables[0], dt.Tables[3]));
                }
            }
            PrepareHTMLend(twWriter);

            twWriter.Write("<br /><br />");

            // this is to display the Totals if totals are presnt for that perticular report. 

            twWriter.Write("</CENTER>");
            twWriter.Write("</BODY>");
            twWriter.Write("</HTML>");

            twWriter.Close();

            #endregion



            //if (!Portfolio.Contains(schReports.ID1))
            //{
            InsertPortfoliodetails("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + schReports.ID1 + ".html", schReports.EmailAddresses);
            //}
            //else
            //{ 

            // updatecustomportfolio("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName  , schReports.ID1);


            //}


        }
        private void PrepareWORDreport(Schedule schReports, DataSet tables, string trimmedstartfromdate, string trimmedstarttodate, string level)
        {
            // TextWriter twWriter = new StreamWriter("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + schReports.ID1 + ".html");
        }
        // This method is to create txt report version
        private void PrepareTEXTreport(Schedule schReports, DataSet tables, string trimmedstartfromdate, string trimmedstarttodate, string level)
        {
            TextWriter twWriter = new StreamWriter("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + schReports.ID1 + ".txt");
            try
            {


                twWriter.Write(" Date range: " + " -" + trimmedstartfromdate + "-" + trimmedstarttodate);
                twWriter.WriteLine();
                preparereport(twWriter, level, tables, " ".PadRight(5));
                InsertPortfoliodetails("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + schReports.ID1 + ".txt", schReports.EmailAddresses);

            }
            catch { }
            finally
            { twWriter.Close(); }




        }
        // This method is to create csv report version
        private void PrepareCSVreport(Schedule schReports, DataSet tables, string trimmedstartfromdate, string trimmedstarttodate, string level)
        {

            TextWriter twWriter = new StreamWriter("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + schReports.ID1 + ".csv");
            try
            {


                twWriter.Write(" Date range: " + " -" + trimmedstartfromdate + "-" + trimmedstarttodate);
                twWriter.WriteLine();

                preparereport(twWriter, level, tables, ",");
                InsertPortfoliodetails("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + schReports.ID1 + ".csv", schReports.EmailAddresses);

            }
            catch { }
            finally
            { twWriter.Close(); }

        }
        // This method is to create PDF report version
        private void PreparePDFreport(Schedule schreports, DataSet dt, string trimmedstartfromdate, string trimmedstarttodate, string level)
        {
            // create a file with the name
            string file = @"c:\\temp\\ " + schreports.Selectedname + " - " + schreports.ReportName + " - " + schreports.ID1 + ".pdf";
            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 5, 5, 10, 10);
            //System.IO.MemoryStream mStream = new System.IO.MemoryStream();            
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(file, FileMode.Create));

            pdfDoc.Open();
            pdfDoc.Add(new iTextSharp.text.Paragraph(schreports.ReportName));

            pdfDoc.Add(new iTextSharp.text.Paragraph(" Date range: " + " -" + trimmedstartfromdate + "-" + trimmedstarttodate));

            pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
            if (level == "4")
            {
                GetPDfTable(pdfDoc, dt.Tables[2], 0);
                pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                int tablevalue = GetLevels(dt);
                if (tablevalue != 1)
                {
                    GetPDfTable(pdfDoc, dt.Tables[GetLevels(dt) + 2], 1);
                }
                else { GetEXTPDFTable(pdfDoc, dt.Tables[0], dt.Tables[3]); }
            }
            else
            {
                if (GetLevels(dt) == 1)
                {
                    GetPDfTable(pdfDoc, dt.Tables[3], 0);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                }
                else if (GetLevels(dt) == 2)
                {
                    GetPDfTable(pdfDoc, dt.Tables[2], 0);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[4], 0);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetEXTPDFTable(pdfDoc, dt.Tables[0], dt.Tables[3]);
                }
                else if (GetLevels(dt) == 3)
                {
                    GetPDfTable(pdfDoc, dt.Tables[2], 0);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[5], 1);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[4], 1);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetEXTPDFTable(pdfDoc, dt.Tables[0], dt.Tables[3]);
                }
                else if (GetLevels(dt) == 4)
                {
                    GetPDfTable(pdfDoc, dt.Tables[2], 0);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[6], 1);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[5], 1);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[4], 1);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetEXTPDFTable(pdfDoc, dt.Tables[0], dt.Tables[3]);
                }
            }

            pdfDoc.Close();
            InsertPortfoliodetails(file, schreports.EmailAddresses);

        }
        public void GetPDfTable(iTextSharp.text.Document pdfDoc, DataTable dataTable, int level)
        {
            int cols = dataTable.Columns.Count;
            int rows = dataTable.Rows.Count;

            PdfPTable pdfTable = new PdfPTable(dataTable.Columns.Count);
            pdfTable.TotalWidth = 100;
            List<string> column = new List<string>();
            //creating header columns
            foreach (DataColumn colu in dataTable.Columns)
            {
                if (colu.ColumnName != "Department")
                {
                    PdfPCell cell = new PdfPCell(new Phrase(colu.ColumnName, new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8, 0)));
                    column.Add(colu.ColumnName);
                    pdfTable.AddCell(cell);
                }
            }
            //creating rows
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (string col in column)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(row[col].ToString().Replace(";", string.Empty), new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8, 0)));
                    pdfTable.AddCell(cell);

                }
            }
            pdfDoc.Add(pdfTable);


        }
        public void GetEXTPDFTable(iTextSharp.text.Document pdfDoc, DataTable des, DataTable dt)
        {
            try
            {





                DataTable table1 = new DataTable();
                DataTable table2 = new DataTable();

                table1.Columns.Clear();
                table2.Columns.Clear();
                var depts = dt.AsEnumerable().Select(s => s.Field<string>("CostCentre")).Distinct();
                foreach (DataColumn colu in des.Columns)
                {

                    if (colu.ColumnName != "Department")
                    {
                        table1.Columns.Add(colu.ColumnName);
                    }
                }
                foreach (DataColumn colu in dt.Columns)
                {

                    table2.Columns.Add(colu.ColumnName);
                }

                foreach (var v in depts)
                {
                    pdfDoc.Add(new iTextSharp.text.Paragraph(v.ToString()));
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    table1.Rows.Clear();
                    table2.Rows.Clear();

                    var dest = from db in des.AsEnumerable()
                               where db.Field<string>("Department") == v.ToString()
                               select new
                               {
                                   Destinationname = db.Field<object>("Destinationname"),
                                   Totalcalls = db.Field<object>("Totalcalls"),
                                   Totalduration = db.Field<object>("Totalduration"),
                                   Cost = db.Field<object>("Cost")

                               };
                    foreach (var destinations in dest)
                    {
                        table1.Rows.Add(destinations.Destinationname.ToString(), destinations.Totalcalls.ToString(), destinations.Totalduration.ToString(), destinations.Cost.ToString());


                    }
                    GetPDfTable(pdfDoc, table1, 0);

                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    var records = from db in dt.AsEnumerable()
                                  where db.Field<string>("CostCentre") == v.ToString()
                                  select new
                                  {
                                      Name = db.Field<object>("Name"),
                                      Extension = db.Field<object>("Extension"),
                                      OutgoingCalls = db.Field<object>("OutgoingCalls"),
                                      OutgoingDuration = db.Field<object>("OutgoingDuration"),
                                      RingResponse = db.Field<object>("RingResponse"),
                                      AbandonedCalls = db.Field<object>("AbandonedCalls"),
                                      IncomingCalls = db.Field<object>("IncomingCalls"),
                                      IncomingDuration = db.Field<object>("IncomingDuration"),
                                      Cost = db.Field<object>("Cost")

                                  };
                    foreach (var all in records)
                    {
                        table2.Rows.Add(all.Name.ToString(), all.Extension.ToString(), all.OutgoingCalls.ToString(), all.RingResponse.ToString(), all.AbandonedCalls.ToString(), all.IncomingCalls.ToString(), all.IncomingDuration.ToString());

                    }

                    GetPDfTable(pdfDoc, table2, 0);

                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                }






            }
            catch { }




        }
        public void GetStringTable(TextWriter twWriter, DataTable dt, int level, string seperator)
        {


            List<string> column = new List<string>();
            //creating table headers
            foreach (DataColumn colu in dt.Columns)
            {
                if (colu.ColumnName != "Department")
                {
                    twWriter.Write(colu.ColumnName);
                    column.Add(colu.ColumnName);
                    twWriter.Write(seperator);
                }
            }
            twWriter.WriteLine();
            //Creating columns
            foreach (DataRow row in dt.Rows)
            {
                foreach (string col in column)
                {
                    twWriter.Write(row[col].ToString().Replace(";", string.Empty));
                    twWriter.Write(seperator);
                }
                twWriter.WriteLine();
            }

        }
        public void GetExtlevelTable(TextWriter twWriter, DataTable des, DataTable dt, string seperator)
        {



            var depts = dt.AsEnumerable().Select(s => s.Field<string>("CostCentre")).Distinct();

            foreach (var v in depts)
            {
                twWriter.WriteLine();
                twWriter.Write(v.ToString());
                twWriter.WriteLine(" ");

                twWriter.WriteLine();

                foreach (DataColumn column in des.Columns)
                {
                    if (column.ColumnName != "Department")
                    {
                        twWriter.Write(column.ColumnName);
                        twWriter.Write(seperator);
                    }

                }


                var dest = from db in des.AsEnumerable()
                           where db.Field<string>("Department") == v.ToString()
                           select new
                           {
                               Destinationname = db.Field<object>("Destinationname"),
                               Totalcalls = db.Field<object>("Totalcalls"),
                               Totalduration = db.Field<object>("Totalduration"),
                               Cost = db.Field<object>("Cost")

                           };
                foreach (var destinations in dest)
                {
                    twWriter.WriteLine();
                    twWriter.Write(destinations.Destinationname.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                    twWriter.Write(destinations.Totalcalls.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                    twWriter.Write(destinations.Totalduration.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                    twWriter.Write(destinations.Cost.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                }
                twWriter.WriteLine();
                foreach (DataColumn column in dt.Columns)
                {
                    if (column.ColumnName != "CostCentre")
                    {
                        twWriter.Write(column.ColumnName);
                        twWriter.Write(seperator);
                    }


                }


                twWriter.WriteLine(" ");
                var records = from db in dt.AsEnumerable()
                              where db.Field<string>("CostCentre") == v.ToString()
                              select new
                              {
                                  Name = db.Field<object>("Name"),
                                  Extension = db.Field<object>("Extension"),
                                  OutgoingCalls = db.Field<object>("OutgoingCalls"),
                                  OutgoingDuration = db.Field<object>("OutgoingDuration"),
                                  RingResponse = db.Field<object>("RingResponse"),
                                  AbandonedCalls = db.Field<object>("AbandonedCalls"),
                                  IncomingCalls = db.Field<object>("IncomingCalls"),
                                  IncomingDuration = db.Field<object>("IncomingDuration"),
                                  Cost = db.Field<object>("Cost")

                              };
                foreach (var all in records)
                {
                    twWriter.WriteLine();
                    twWriter.Write(all.Name.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                    twWriter.Write(all.Extension.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                    twWriter.Write(all.RingResponse.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                    twWriter.Write(all.IncomingCalls.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                    twWriter.Write(all.IncomingDuration.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                    twWriter.Write(all.AbandonedCalls.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                    twWriter.Write(all.OutgoingCalls.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                    twWriter.Write(all.OutgoingDuration.ToString().Replace(" ", string.Empty));
                    twWriter.Write(seperator);
                    twWriter.Write(all.Cost.ToString().Replace(" ", string.Empty));

                }

                twWriter.WriteLine(" ");


            }



        }
        // this method will determine the level of reports
        public void preparereport(TextWriter twWriter, string level, DataSet dt, string seperator)
        {
            if (level == "4")
            {
                GetStringTable(twWriter, dt.Tables[2], 0, seperator);
                twWriter.WriteLine();
                int tablevalue = GetLevels(dt);
                if (tablevalue != 1)
                {
                    GetStringTable(twWriter, dt.Tables[GetLevels(dt) + 2], 1, seperator);
                }
                else { GetExtlevelTable(twWriter, dt.Tables[0], dt.Tables[3], seperator); }
            }
            else
            {
                if (GetLevels(dt) == 1)
                {
                    GetStringTable(twWriter, dt.Tables[3], 0, seperator);
                    twWriter.WriteLine();
                }
                else if (GetLevels(dt) == 2)
                {
                    GetStringTable(twWriter, dt.Tables[2], 0, seperator);
                    twWriter.WriteLine();
                    GetStringTable(twWriter, dt.Tables[4], 0, seperator);
                    twWriter.WriteLine();
                    GetExtlevelTable(twWriter, dt.Tables[0], dt.Tables[3], seperator);
                }
                else if (GetLevels(dt) == 3)
                {
                    GetStringTable(twWriter, dt.Tables[2], 0, seperator);
                    twWriter.WriteLine();
                    GetStringTable(twWriter, dt.Tables[5], 1, seperator);
                    twWriter.WriteLine();
                    GetStringTable(twWriter, dt.Tables[4], 1, seperator);
                    twWriter.WriteLine();
                    GetExtlevelTable(twWriter, dt.Tables[0], dt.Tables[3], seperator);
                }
                else if (GetLevels(dt) == 4)
                {
                    GetStringTable(twWriter, dt.Tables[2], 0, seperator);
                    twWriter.WriteLine();
                    GetStringTable(twWriter, dt.Tables[6], 1, seperator);
                    twWriter.WriteLine();
                    GetStringTable(twWriter, dt.Tables[5], 1, seperator);
                    twWriter.WriteLine();
                    GetStringTable(twWriter, dt.Tables[4], 1, seperator);
                    twWriter.WriteLine();
                    GetExtlevelTable(twWriter, dt.Tables[0], dt.Tables[3], seperator);
                }
            }
        }
        # endregion

       
        */

        #endregion -------------------------------------------------------------------------------------------------------------------



        public Form1()
        {
            GenerateReports.LogMessageToFile(string.Empty);
            GenerateReports.LogMessageToFile("Starting");
            InitializeComponent();
            //        SendEmail.sendEmails();
            Datamethods.Starttime = DateTime.Now.TimeOfDay.Hours;
            //SendEmail.SendErrormessage();

            var methodName = this.GetType().FullName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name;

            try
            {

                // This loop is for report frequency
                foreach (tbl_reportfrequency freq in Datamethods.Frequency())
                {
                    GenerateReports.LogMessageToFile(freq.Frequency_Name);

                    // this loop is to get all reports for the report frequency....
                    foreach (Schedule report in Datamethods.GetReportsData(freq.Frequency_Name.ToString()))
                    {
                        try
                        {
                            switch (report.Type)
                            {
                                case "html":
                                    //Datamethods.Undoreport(report);
                                    GenerateReports.GenerateHTMLReport(report);
                                    Datamethods.Updatereport(report);
                                    break;
                                case "word":
                                    //  Datamethods.Undoreport(report);
                                    GenerateReports.GenerateWORDReport(report);
                                    Datamethods.Updatereport(report);
                                    break;
                                case "text":
                                    //  Datamethods.Undoreport(report);
                                    GenerateReports.GenerateTextReport(report);
                                    Datamethods.Updatereport(report);
                                    break;
                                case "PDF":
                                    //Datamethods.Undoreport(report);
                                    GenerateReports.GeneratePDFReport(report);
                                    Datamethods.Updatereport(report);
                                    break;
                                case "xml":
                                    //  Datamethods.Undoreport(report);
                                    GenerateReports.GenerateXMLReport(report);
                                    Datamethods.Updatereport(report);
                                    break;
                                case "Excel":
                                    GenerateReports.GenerateExcelReport(report);
                                    Datamethods.Updatereport(report);
                                    break;
                                default:
                               //     Datamethods.Undoreport(report);
                                    GenerateReports.GenerateCSVReport(report);
                                    Datamethods.Updatereport(report);
                                    break;


                            }
                        }
                        catch { }
                        finally { }
                        // Need to implement i-dispose interface here  to make GC to go for collection......... 
                        Dispose(true);
                    }


                    if (freq.Frequency_Name == "hourly" || freq.Frequency_Name == "fixed")
                    {
                        try
                        {
                            Datamethods.Updatetimes(freq);
                        }
                        catch { }
                    }
                }

                SendEmail.sendEmails();
           //     SendEmail.SendErrormessage();

                //  Datamethods.UpdateReportStatus(status);
            }
            catch(Exception ex)
            {
                GenerateReports.LogMessageToFile(methodName, ex);
            }

            GenerateReports.LogMessageToFile("Ending");
            Environment.Exit(1);


            /* ...........................................................Danger STOP............................................................
               -----------------------------------------------------Do not proceeed the code below is disgusting ----------------------------------------------------
             ...............................................................................................................................
            
             */
            # region------------- old code -------------------------

            /*       CreateDailySchedules();

            List<int> Portfolio = Getportfolioreports();
            SqlConnection oConn = new SqlConnection(TEMConnectionString);
            oConn.Open();

            SqlCommand oCmd = new SqlCommand("spProcessDailySchedules_test", oConn); //returns all the data needed to run the report
            oCmd.CommandType = CommandType.StoredProcedure;
            SqlDataReader drReader = oCmd.ExecuteReader();
            ArrayList alDailySchedules = new ArrayList();
            ArrayList alExport = new ArrayList();
         
            drReader.Read();

            do
            {
                if (drReader.HasRows == true)
                {
                    //TODO; edit schedule class for new column property
                    Schedule SchedulesToday = new Schedule();

                    SchedulesToday.ID = drReader.GetInt32(0); //unique id for each schedule
                    SchedulesToday.StoredProcedureName = drReader.GetString(1); //sp name
                    SchedulesToday.ReportName = drReader.GetString(2); //name of the report
                    SchedulesToday.CreatedDate = drReader.GetDateTime(3); //create datetime of the schedule
                    SchedulesToday.ListofFilters = drReader.GetString(4); //csv list of all filters - prob needs seperated based on @ symbol
                    SchedulesToday.Frequency = drReader.GetString(5); //frequency - how often report is to be delivered (daily,weekly,monthly)
                    SchedulesToday.Type = drReader.GetString(6); //type - csv,html,word etc
                    SchedulesToday.EmailAddresses = drReader.GetString(7); //csv list of all email address who need to receive the reports
                    SchedulesToday.Columns = drReader.GetString(8); //column information regarding the stored procedures
                   
                   if (!drReader.IsDBNull(9))
                   {
                       SchedulesToday.GraphBindings = drReader.GetString(9); //the x and y axis which need to be bound to the report
                   }
                   //else { SchedulesToday.GraphBindings = 0,0 }
                   SchedulesToday.ID1 = drReader.GetInt32(10);
                    SchedulesToday.Chosennodelist = drReader.GetString(11);
                   // string nodelist =
                    SchedulesToday.Selectedname = drReader.GetString(12);

                    if (!drReader.IsDBNull(13))
                    {
                        SchedulesToday.Totals = drReader.GetString(13);
                    }
                    if (!drReader.IsDBNull(14))
                    {
                        SchedulesToday.Dropdownid = drReader.GetInt32(14);
                    }
                    alDailySchedules.Add(SchedulesToday);
         
                }

            }
            while (drReader.Read());

            drReader.Close();
            oConn.Close();

            foreach (Schedule schReports in alDailySchedules)
            {

                if (schReports.StoredProcedureName.Contains("Mobile"))
                {

                    alExport = GetReportdata(schReports.StoredProcedureName, schReports.ListofFilters, schReports.Columns, schReports.Chosennodelist);

                }
                else
                {
                    //TODO; call a function which runs the report and generates the cdr results - might need to add total column


                    alExport = GenerateCDRResults(schReports.StoredProcedureName, schReports.ListofFilters, schReports.EmailAddresses, schReports.Columns, schReports.Chosennodelist, schReports.ID1, schReports.Selectedname, schReports.Totals);

                }


                string[] GraphBindXY = schReports.GraphBindings.Split(',');
                string Xaxis;
                string Yaxis;
                bool blnHaveChart = false;

                if (GraphBindXY[0] == "")
                {
                    Xaxis = "";
                    Yaxis = "";

                }
                else
                {
                    Xaxis = GraphBindXY[0];
                    Yaxis = GraphBindXY[1];
                    blnHaveChart = true;
                }
                ReportHasData = true;
                if (ReportHasData == true)
                {
                    string strExtension;

                    switch (schReports.Type)
                    {
                        case "html":
                            strExtension = ".html";
                            break;
                        case "word":
                            strExtension = ".doc";
                            break;
                        case "text":
                            strExtension = ".txt";
                            break;
                        case "PDF":
                            strExtension = ".PDF";
                            break;
                        default:

                            strExtension = ".csv";
                            break;
                    }

                    DateTime dtImageCreated = new DateTime();
                    dtImageCreated = DateTime.Now;
                    string strUniqueDateTime = String.Format("{0:ddMMyyyyHHmmss}", dtImageCreated);


                    #region Gathering the required data,Preparing the mail message and sending  ---------------
                    try
                    {


                        if (schReports.StoredProcedureName.Contains("spSelectDepartmentalBreakdownreportLevel"))
                        {

                            GenerateCostcenterreport(schReports);

                        }
                        else
                        {
                            TextWriter twWriter = new StreamWriter("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + strUniqueDateTime + "-" + schReports.ID1 + strExtension);

                            //pass the cdr results into the correct type)
                            string[] tofromdates1 = schReports.ListofFilters.Split('=', ',');
                            switch (schReports.Type)
                            {

                                #region HTML stuff --------------------------------------------------------------
                                case "html": //html
                                    #region "HTML"


                                    string trimmedstartfromdate = tofromdates1[5].TrimStart('\'');
                                    string trimmedfromdate1 = trimmedstartfromdate.TrimEnd('\'');

                                    string trimmedstarttodate = tofromdates1[7].TrimStart('\'');
                                    string trimmedtodate1 = trimmedstarttodate.TrimEnd('\'');

                                    twWriter.Write("<HTML>");
                                    twWriter.Write("<HEAD>");
                                    twWriter.Write("<TITLE>");
                                    twWriter.Write(schReports.ReportName);
                                    twWriter.Write("</TITLE>");

                                    twWriter.Write("<style type='text/css'>");
                                    twWriter.Write(".style1");
                                    twWriter.Write("{");
                                    twWriter.Write("font-size: 70pt;");
                                    twWriter.Write("}");
                                    twWriter.Write("</style>");
                                    twWriter.Write("</HEAD>");
                                    twWriter.Write("<BODY BGCOLOR='white'>");
                                    twWriter.Write("<CENTER>");

                                    twWriter.Write("<img src='http://goo.gl/3NmNM' />");
                                    twWriter.Write("<BR><BR><H2>" + schReports.ReportName + "</H2>");

                                    #region may need removed from code if table does not work
                                    twWriter.Write("<CENTER>");
                                    twWriter.Write("<table cellpadding='0' cellspacing='0' border='0' class='opaque'>");
                                    twWriter.Write("<TR>");
                                    twWriter.Write("<TD>");
                                    twWriter.Write("</TD>");
                                    twWriter.Write("<TD>");
                                    twWriter.Write("</TD>");
                                    twWriter.Write("<TD align ='left'>");
                                    twWriter.Write("Date Range :  ");
                                    twWriter.Write(trimmedstartfromdate);

                                    twWriter.Write("  --  ");
                                    twWriter.Write(trimmedstarttodate);
                                    twWriter.Write("</TD>");
                                    twWriter.Write("<TD>");
                                    twWriter.Write("</TD>");
                                    twWriter.Write("</TR>");
                                    twWriter.Write("<CENTER>");
                                    twWriter.Write("<tr>");
                                    twWriter.Write("<td style='width: 12px; height: 1px;'><img alt='' src='http://goo.gl/1ERss' width='12' height='1' /></td>");
                                    twWriter.Write("<td style='width: 18px; height: 16px; background-image: url(http://goo.gl/F1Y6K);'><img alt='' src='http://goo.gl/1ERss' width='18' height='16' /></td>");
                                    twWriter.Write("<td style='width: 910px; height: 16px; background-image: url(http://goo.gl/twveK);'><img alt='' src='http://goo.gl/1ERss' width='910' height='16' /></td>");
                                    twWriter.Write("<td style='width: 23px; height: 16px; background-image: url(http://goo.gl/kCStA);'><img alt='' src='http://goo.gl/1ERss' width='23' height='16' /></td>");
                                    twWriter.Write("</tr>");
                                    twWriter.Write("<tr>");
                                    twWriter.Write("<td style='width: 12px; height: 1px;'><img alt='' src='http://goo.gl/1ERss' width='12' height='1' /></td>");
                                    twWriter.Write("<td style='width: 18px; background-image: url(http://goo.gl/sJxcZ);'></td>");
                                    twWriter.Write("<td>");

                                    twWriter.Write("<table WIDTH='100%' cellpadding='5' cellspacing='0'align ='center'>");
                                    twWriter.Write("<CENTER>");
                                    twWriter.Write("<TR bgcolor= '#CAD8EC'>");

                                    twWriter.Write("</TR>");
                                    twWriter.Write("<TR bgcolor= '#CAD8EC'><TD>");

                                    oConn.Open();

                                    String origChosenNode = schReports.Chosennodelist;
                                    string[] choosennode1 = schReports.Chosennodelist.Split('=');
                                    string param = "'" + choosennode1[1] + "'";
                                    string name = choosennode1[0] + "=";


                                    schReports.Chosennodelist = name + param;

                                    SqlCommand oCmdCDR = new SqlCommand("exec " + schReports.StoredProcedureName + " " + schReports.ListofFilters + "," + schReports.Chosennodelist, oConn);
                                    //oConn.ConnectionTimeout = 3000;
                                    oCmdCDR.CommandTimeout = 3000;

                                    ArrayList FILTERING = new ArrayList();
                                    string[] Columns = schReports.Columns.Split(',');
                                    ArrayList nameofreport = new ArrayList();
                                    ArrayList value = new ArrayList();

                                    //string[] totals = "";
                                    if (schReports.Totals != "")
                                    {

                                        string totalstrimstat = schReports.Totals.TrimStart(',');
                                        string totalstrimstatEnd = totalstrimstat.TrimEnd(',');
                                        string[] totals = totalstrimstatEnd.Split(',');
                                        for (int i = 0; i < totals.Length; i++)
                                        {
                                            // string[] totals = totalstrimstatEnd.Split(',');
                                            // nameofreport.Add(totals[0]);
                                            //int i = 0;
                                            nameofreport.Add(totals[i]);
                                            // i++;

                                        }
                                    }



                                    //TODO; change to total columns
                                    foreach (string Column in Columns)
                                    {
                                        string[] arrColumnNameAndLength = Column.Split('@', '#');

                                        //do function and pass in column name and amount
                                        FILTERING.Add(arrColumnNameAndLength[0]);
                                    }

                                    // this is to Get the nxt result of datareader for totals.......
                                    if (nameofreport.Count >= 1)
                                    {
                                        SqlDataReader drReader1 = oCmdCDR.ExecuteReader();
                                        //value.Add(drReader1.FieldCount);
                                        drReader1.NextResult();

                                        while (drReader1.Read())
                                        {
                                            foreach (string total in nameofreport)
                                            {



                                                value.Add(drReader1[total]);


                                            }



                                        }
                                        drReader1.Close();

                                    }

                                    nameofreport.Add("Count");
                                    value.Add(Callcount.ToString());




                                    if (Xaxis != "" && Yaxis != "")
                                    {
                                        CreateBarChart(oCmdCDR, drReader, strUniqueDateTime, Xaxis, Yaxis, schReports.ReportName);
                                        CreatePieChart(oCmdCDR, drReader, strUniqueDateTime, Xaxis, Yaxis, schReports.ReportName);
                                    }
                                    oConn.Close();

                                    //TODO; pass in totalcolumns arraylist
                                    twWriter.Write(GenerateHTMLCDRResults(schReports.StoredProcedureName, schReports.ListofFilters, schReports.EmailAddresses, schReports.Columns, schReports.Chosennodelist, schReports.ID1, nameofreport));


                                    twWriter.Write("</TD></TR>");
                                    twWriter.Write("</table >");
                                    twWriter.Write("</td>");
                                    twWriter.Write("<td style='width: 23px; background-image: url(http://goo.gl/HQSbi);'></td>");
                                    twWriter.Write("</tr>");
                                    twWriter.Write("<tr>");
                                    twWriter.Write("<td style='width: 12px; height: 1px;'><img alt='' src='http://goo.gl/1ERss' width='12' height='1' /></td>");
                                    twWriter.Write("<td style='width: 18px; height: 23px; background-image: url(http://goo.gl/aZ7T4);'><img alt='' src='http://goo.gl/1ERss' width='18' height='23' /></td>");
                                    twWriter.Write("<td style='width: 910px; height: 23px; background-image: url(http://goo.gl/Ljz7E);'><img alt='' src='http://goo.gl/1ERss' width='910' height='23' /></td>");
                                    twWriter.Write("<td style='width: 23px; height: 23px; background-image: url(http://goo.gl/caJ7u);'><img alt='' src='http://www.sentelcallmanagerpro.com/images/spacer.gif' width='23' height='23' /></td>");
                                    twWriter.Write("</tr>");
                                    twWriter.Write("</table>");

                                    #endregion


                                    twWriter.Write("</table>");
                                    twWriter.Write("<br /><br /><br /><br />");



                                    // this is to display the Totals if totals are presnt for that perticular report. 
                                    if (nameofreport.Count >= 1)
                                    {
                                        twWriter.Write("<table>");
                                        twWriter.Write("<tr>");

                                        for (int i = 0; i < nameofreport.Count; i++)
                                        {

                                            twWriter.Write("<td style='width:120'>");
                                            twWriter.Write(@nameofreport[i].ToString() + " " + "</td>");

                                            // twWriter.Write("<td>&nbsp,&nbsp</td>");
                                            //twWriter.Write("</td>");

                                        }
                                        twWriter.Write("</tr>");
                                        twWriter.Write("<tr>");
                                        for (int j = 0; j < value.Count; j++)
                                        {
                                            twWriter.Write("<td style='width:50'>");
                                            twWriter.Write(" " + value[j].ToString() + "," + "</td>");
                                        }


                                        twWriter.Write("</table>");
                                    }


                                    if (blnHaveChart == true)
                                    {
                                        twWriter.Write("<table><tr><td>");


                                        twWriter.Write("<img src='c:\\visualstudio\\projects\\TemAutomatedReports\\images\\chart" + strUniqueDateTime + ".png'></td>");
                                        twWriter.Write("<td>&nbsp</td>");
                                        twWriter.Write("<td><img src='c:\\visualstudio\\projects\\TemAutomatedReports\\images\\pie" + strUniqueDateTime + ".png'></td></tr>");
                                        twWriter.Write("</table>");
                                    }

                                    twWriter.Write("</CENTER>");
                                    twWriter.Write("</BODY>");
                                    twWriter.Write("</HTML>");

                                    twWriter.Close();

                                    #endregion



                                    if (!Portfolio.Contains(schReports.ID1))
                                    {
                                        InsertPortfoliodetails("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + strUniqueDateTime + "-" + schReports.ID1 + strExtension, schReports.EmailAddresses);
                                    }
                                    else
                                    {

                                        updatecustomportfolio("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + strUniqueDateTime + "-" + schReports.ID1 + strExtension, schReports.ID1);


                                    }
                                    // CreateMailMessage(schReports.EmailAddresses, schReports.ReportName, strExtension, strUniqueDateTime, blnHaveChart, schReports.Selectedname);

                                    break;
                                #endregion
                                #region word stuff --------------------------------------------------------------

                                case "word": //word
                                    /*
                               Some languages support optional arguments.  VB.NET is one of them, it is also used often in 
                               COM programming with Office.  C# is not one of them.  It has to explicitly pass Missing to
                               indicate that the optional argument is in fact omitted.
                               */
            /*    object oMissing = System.Reflection.Missing.Value;

                object oEndOfDoc = "\\endofdoc";

                //Start word and create a new document.
                Word._Application oWord;
                Word._Document oDoc;

                oWord = new Word.Application();
                oWord.Visible = true;
                //oWord.Visible = false;

                //see above for definition
                oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                //Insert a paragraph at the beginning of the document.
                Word.Paragraph oPara1;
                oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
                oPara1.Range.Text = "Heading 1";
                oPara1.Range.Font.Bold = 1;
                oPara1.Format.SpaceAfter = 24; //24 pt spacing after the paragraph
                oPara1.Range.InsertParagraphAfter();

                //Insert a paragraph at the end of the document.
                Word.Paragraph oPara2;
                object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
                oPara2.Range.Text = "Heading 2";
                oPara2.Format.SpaceAfter = 6;
                oPara2.Range.InsertParagraphAfter();

                //Insert another paragraph.
                Word.Paragraph oPara3;
                oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
                oPara3.Range.Text = "This is a sentence of normal text. Now here is a table:";
                oPara3.Range.Font.Bold = 0;
                oPara3.Format.SpaceAfter = 24;
                oPara3.Range.InsertParagraphAfter();

                //Insert a 3 * 5 table, fill it with data and make the first row bold and italic
                Word.Table oTable;
                Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oTable = oDoc.Tables.Add(wrdRng, 3, 5, ref oMissing, ref oMissing);
                oTable.Range.ParagraphFormat.SpaceAfter = 6;
                int row, column;
                string strText;

                for (row = 1; row <= 3; row++)
                {
                    for (column = 1; column <= 5; column++)
                    {
                        strText = "r" + row + "c" + column;
                        oTable.Cell(row, column).Range.Text = strText;
                    }
                }
                oTable.Rows[1].Range.Font.Bold = 1;
                oTable.Rows[1].Range.Font.Italic = 1;

                //Add some text after the table

                Word.Paragraph oPara4;
                oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
                oPara4.Range.InsertParagraphBefore();
                oPara4.Range.Text = "And here's another table:";
                oPara4.Format.SpaceAfter = 24;
                oPara4.Range.InsertParagraphAfter();

                //Insert a 5 * 2 table, fill it with data, and change the columns widths.
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oTable = oDoc.Tables.Add(wrdRng, 5, 2, ref oMissing, ref oMissing);
                oTable.Range.ParagraphFormat.SpaceAfter = 6;

                for (row = 1; row <= 5; row++)
                {
                    for (column = 1; column <= 2; column++)
                    {
                        strText = "r" + row + "c" + column;
                        oTable.Cell(row, column).Range.Text = strText;
                    }
                }
                oTable.Columns[1].Width = oWord.InchesToPoints(2); //change width of columns 1 and 2
                oTable.Columns[2].Width = oWord.InchesToPoints(3);

                //keep inserting text. when you get to 7 inches from the top of the
                //document, insert a hard break point
                object oPos;
                double dPos = oWord.InchesToPoints(7);
                oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertParagraphAfter();
                do
                {
                    wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    wrdRng.ParagraphFormat.SpaceAfter = 6;
                    wrdRng.InsertAfter("A line of text");
                    wrdRng.InsertParagraphAfter();
                    oPos = wrdRng.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage);
                }
                while (dPos >= Convert.ToDouble(oPos));
                object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                object oPageBreak = Word.WdBreakType.wdPageBreak;
                wrdRng.Collapse(ref oCollapseEnd);
                wrdRng.InsertBreak(ref oPageBreak);
                wrdRng.Collapse(ref oCollapseEnd);
                wrdRng.InsertAfter("We're now on page 2. Here's my chart:");
                wrdRng.InsertParagraphAfter();

                //Insert a chart
                Word.InlineShape oShape;
                object oClassType = "MSGraph.Chart.8"; //might need changed to 8 - think it depends of version of ms
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                //Demonstrate use of late bound oChart and oChartApp objects to
                //manipulate the chart object with MSGraph.
                object oChart;
                object oChartApp;
                oChart = oShape.OLEFormat.Object;
                oChartApp = oChart.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, oChart, null);

                //change the chart type to line
                object[] Parameters = new object[1];
                Parameters[0] = 4; //xlline = 4
                oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty, null, oChart, Parameters);

                //Update the chart image and quit MSGraph.
                oChartApp.GetType().InvokeMember("Update", BindingFlags.InvokeMethod, null, oChartApp, null);
                oChartApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, oChartApp, null);

                //you can make further changes from the charts from here

                //set the width of the chart
                oShape.Width = oWord.InchesToPoints(6.25f);
                oShape.Height = oWord.InchesToPoints(3.75f);

                //add some text to the charts
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng.InsertParagraphAfter();
                wrdRng.InsertAfter("The End");

                object FilePath = "c:\\" + schReports.ReportName + strExtension;
                object FileFormat = Word.WdSaveFormat.wdFormatDocument;
                object LockComments = false;
                object AddToRecentFiles = false;
                object ReadOnlyRecommended = false;
                object EmbedTrueTypeFonts = false;
                object SaveNativePictureFormat = true;
                object SaveFormsData = true;
                object SaveAsAOCELetter = false;
                object Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUSASCII;
                object InsertLineBreaks = false;
                object AllowSubstitutions = false;
                object LineEnding = Word.WdLineEndingType.wdCRLF;
                object AddBiDiMarks = false;

                object saveChanges = false;
                object doSaveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;



                oDoc.Close(ref doSaveChanges, ref oMissing, ref oMissing);

                //oDoc.SaveAs(ref FilePath, ref FileFormat, ref LockComments,
                //ref oMissing, ref AddToRecentFiles, ref oMissing,
                //ref ReadOnlyRecommended, ref EmbedTrueTypeFonts,
                //ref SaveNativePictureFormat, ref SaveFormsData,
                //ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks,
                //ref AllowSubstitutions, ref LineEnding, ref AddBiDiMarks);





                oWord.Quit(ref saveChanges, ref oMissing, ref oMissing);
                //oWord.ActiveDocument.Close(ref doSaveChanges, ref oMissing, ref oMissing);

                //Create a mail message for the schedule report with attachement
                //CreateMailMessage(schReports.EmailAddresses, schReports.ReportName, strExtension, strUniqueDateTime, schReports.Selectedname);
                // TO insert portfolio report.......
                if (!Portfolio.Contains(schReports.ID1))
                {
                    InsertPortfoliodetails("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + strUniqueDateTime + "-" + schReports.ID1 + strExtension, schReports.EmailAddresses);

                }
                else
                {

                    updatecustomportfolio("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + strUniqueDateTime + "-" + schReports.ID1 + strExtension, schReports.ID1);

                }
                break;
            #endregion
            #region Text stuff --------------------------------------------------------------
            case "Text":
                for (int i = 0; i < alExport.Count; i++)
                {
                    if (alExport[i].ToString() == "\r\n")
                        twWriter.Write(alExport[i].ToString());
                    else
                        twWriter.Write(alExport[i].ToString() + '\t');
                }

                twWriter.Close();

                //Create a mail message for the schedule report with attachment
                //CreateMailMessage(schReports.EmailAddresses, schReports.ReportName, strExtension, strUniqueDateTime, schReports.Selectedname);
                // TO insert portfolio report.......
                if (!Portfolio.Contains(schReports.ID1))
                {
                    InsertPortfoliodetails("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + strUniqueDateTime + "-" + schReports.ID1 + strExtension, schReports.EmailAddresses);
                }
                else
                {
                    updatecustomportfolio("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + strUniqueDateTime + "-" + schReports.ID1 + strExtension, schReports.ID1);


                }
                break;
            #endregion
            #region PDF stuff ---------------------------------------------------------------
            case "PDF":
                //create a new document
                Doc.Document dDoc = CreateDocuments(schReports.ReportName, alExport, tofromdates1[5], tofromdates1[7]);
                int id1 = schReports.ID1;
                string strReport = schReports.ReportName;
                DateTime strDateTime = schReports.CreatedDate;


                PdfDocumentRenderer pdfrender = new PdfDocumentRenderer();

                pdfrender.Document = dDoc;

                pdfrender.RenderDocument();

                string strFileName = "c:\\temp\\ " + schReports.Selectedname + " - " + id1 + ".pdf";

                pdfrender.Save(strFileName);

                if (!Portfolio.Contains(schReports.ID1))
                {
                    InsertPortfoliodetails(strFileName, schReports.EmailAddresses);

                }
                else
                {
                    updatecustomportfolio(strFileName, schReports.ID1);
                }

                break;
            #endregion
            # region CSV stuff----------------------------------------------------------------
            default: //csv
                for (int i = 0; i < alExport.Count; i++)
                {
                    if (alExport[i].ToString() == "\r\n")
                        twWriter.Write(alExport[i].ToString());
                    else
                        twWriter.Write(alExport[i].ToString() + ',');
                }

                twWriter.Close();
                // TO insert portfolio report.......
                if (!Portfolio.Contains(schReports.ID1))
                {
                    InsertPortfoliodetails("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + strUniqueDateTime + "-" + schReports.ID1 + strExtension, schReports.EmailAddresses);
                }
                else
                {
                    updatecustomportfolio("c:\\temp\\" + schReports.Selectedname + " - " + schReports.ReportName + strUniqueDateTime + "-" + schReports.ID1 + strExtension, schReports.ID1);

                }
                //Create a mail message for the schedule report with attachement
                // CreateMailMessage(schReports.EmailAddresses, schReports.ReportName, strExtension, strUniqueDateTime, schReports.Selectedname);

                break;
            #endregion
        }

    }

#endregion
}

catch
{
    oConn.Close();
    CreateErrorMessage(schReports.EmailAddresses, schReports.ID1, schReports.CreatedDate, schReports.ReportName, schReports.Selectedname);
    // break;
}
}
else
{
//Creates a text file for each blank report. needs extended to identify each blank schedule
//  TextWriter twWriter = new StreamWriter("c:\\TemautomatedReportslogfile.txt");
// TextWriter twWriter = new StreamWriter("c:\\temp\\" + schReports.ReportName);

//BlankReportLog(schReports.ID1, schReports.EmailAddresses, schReports.CreatedDate, schReports.ReportName);
CreateErrorMessage(schReports.EmailAddresses, schReports.ID1, schReports.CreatedDate, schReports.ReportName);


}

}
//Creates a message to servicedelivery and sends the blank report.txt file - move into above else after testing
// CreateErrorMessage();

//LOOP THROUGH ALL PROCESSED SCHEDULES
sendEmails();
GC.Collect();
// sendportfolioreports();



#region Updating the filter--------------------------------------------------------------
foreach (Schedule schReports in alDailySchedules)
{
//TAKE LISTOFFILTERS FROM THE SCHEDULE AND SPLIT IT TO SEPERATE THE FROM AND TO DATE FROM THE OTHER FILTERS
try
{

// ArrayList alExport = GenerateCDRResults(schReports.StoredProcedureName, schReports.ListofFilters, schReports.EmailAddresses, schReports.Frequency, schReports.Chosennodelist, schReports.ID1, schReports.Selectedname, schReports.Totals);

//string[] tofromdates1 = schReports.ListofFilters.IndexOf(tofromdates1);
string[] tofromdates = schReports.ListofFilters.Split('=', ',');

string trimmedstartfromdate = tofromdates[5].TrimStart('\'');
string trimmedfromdate = trimmedstartfromdate.TrimEnd('\'');

string trimmedstarttodate = tofromdates[7].TrimStart('\'');
string trimmedtodate = trimmedstarttodate.TrimEnd('\'');

//DateTime fromdate,todate;
string strMyfromDateTime = trimmedfromdate;
string strMytoDateTime = trimmedtodate;

System.IFormatProvider format = new System.Globalization.CultureInfo("en-US", true);
string expectedformat = "yyyy-MM-dd HH:mm:ss";


System.DateTime fromdate = System.DateTime.ParseExact(strMyfromDateTime, expectedformat, format);
System.DateTime todate = System.DateTime.ParseExact(strMytoDateTime, expectedformat, format);
//string parameter2;
//string parameter1;

switch (schReports.Frequency)
{
    case "daily":
                          
        todate = todate.AddDays(1);
        if (schReports.Dropdownid != 6)
        { fromdate = fromdate.AddDays(1); }
        else {if (todate.Date.ToString("dd") == "01") { fromdate.AddMonths(1); }}
                               
                            
         string fromdate1 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", fromdate);
         string todate1 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", todate);

            fromdate1 = "'" + fromdate1 + "'";
            todate1 = "'" + todate1 + "'";
         string parameter1 = schReports.ListofFilters.Replace(tofromdates[5], fromdate1);
            string parameter2 = parameter1.Replace(tofromdates[7], todate1);
        schReports.ListofFilters = parameter2;

        // tofromdates.GetValue(5);
        //tofromdates.GetValue(7);


        break;
    case "weekly":
        fromdate = fromdate.AddDays(7);
        todate = todate.AddDays(7);
        fromdate1 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", fromdate);
        todate1 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", todate);

        fromdate1 = "'" + fromdate1 + "'";
        todate1 = "'" + todate1 + "'";

        //String.Format("{0:MM/dd/yyyy}", dt);  


        //string  from1date = fromdate.ToString();  
        parameter1 = schReports.ListofFilters.Replace(tofromdates[5], fromdate1);
        parameter2 = parameter1.Replace(tofromdates[7], todate1);
        schReports.ListofFilters = parameter2;

        break;
    case "monthly":
        fromdate = fromdate.AddMonths(1);
        todate = todate.AddMonths(1);
        fromdate1 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", fromdate);
        todate1 = string.Format("{0:yyyy-MM-dd HH:mm:ss}", todate);

        fromdate1 = "'" + fromdate1 + "'";
        todate1 = "'" + todate1 + "'";

        parameter1 = schReports.ListofFilters.Replace(tofromdates[5], fromdate1);
        parameter2 = parameter1.Replace(tofromdates[7], todate1);
        schReports.ListofFilters = parameter2;
        break;
}

//string newparameter = parameter2;
SqlConnection oConn3 = new SqlConnection(TEMConnectionString);
oConn3.Open();
SqlParameter p1;
SqlCommand Cmd = new SqlCommand("spupdatefilterdatefields", oConn3);

Cmd.CommandType = CommandType.StoredProcedure;

p1 = new SqlParameter("@filter", SqlDbType.VarChar, 2000);
p1.Value = schReports.ListofFilters;
Cmd.Parameters.Add(p1);

p1 = new SqlParameter("@id", SqlDbType.Int);
p1.Value = schReports.ID1;
Cmd.Parameters.Add(p1);
Cmd.ExecuteNonQuery();
oConn3.Close();
//switch case statement - iofs daily add a day to both from to dates
//if weekely add 7 days to from andf to dates etc etc

//CALL AN SP TO UPDATE THE FROMDATE AND TODATE DEPENDENT ON WHETHER ITS DAILY, WEEKLY, MONTHLY

//sp shoudl be update tbl_automated _reports where
//schedule name = schreports.reportname 
//and 


Updatefilters(schReports.ListofFilters, schReports.ID1);
}
catch 
{

}
#endregion

}
 
//Deletes all processed reports for that day


            

DeleteProcessedReports();
tbProgress.Text = "Emails sent successfully";

//Environment.Exit(1);*/
            #endregion

        }


    }
}
