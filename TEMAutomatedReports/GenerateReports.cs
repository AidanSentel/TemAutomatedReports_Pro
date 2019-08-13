using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using System.Collections;
using iTextSharp.text.pdf;
using iTextSharp.text;
namespace TEMAutomatedReports
{

    class GenerateReports
    {
        public static DateTime Fromdate { get; set; }
        public static DateTime Todate { get; set; }
        // This method will create HTML file and inserts into portfolio reports.....
        public static void GenerateHTMLReport(Schedule report)
        {
            // Getting report data
            DataSet tables = ExtecuteReport(report);
            if (tables.Tables[0].Rows.Count >= 1 || report.StoredProcedureName.Contains("spSelectDepartmentalBreakdownreportLevel"))
            {

                TextWriter twWriter = new StreamWriter("c:\\temp\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".html");
                string[] filt = report.ListofFilters.Split('=', ',');

                try
                {
                    twWriter.Write("<meta http-equiv='Content-Type' content='text/html'; charset='utf-8'>");
                    twWriter.Write("<script language='javascript' type='text/javascript' src='https://www.sentelsolutions.com/Scripts/jquery-1.4.2.min.js'></script>");
                    twWriter.Write("<link href='https://www.sentelsolutions.com/bootstrap/css/bootstrap.css' rel='stylesheet' />");
                    twWriter.Write("<script src='https://www.sentelsolutions.com/bootstrap/Scripts/bootstrap.min.js'></script>");
                    if (Otherreports(report.ReportingSection) == false)
                    {
                        twWriter.Write("</br>");
                        twWriter.Write("</br>");
                        twWriter.Write("<CENTER><h1>" + report.ReportName + "</h1></CENTER>");
                        twWriter.Write("</br>");
                        twWriter.Write("</br>");
                        twWriter.Write("<b>Schedule Name: " + " " + report.Selectedname + "</b>");
                        twWriter.Write("</br>");
                        twWriter.Write("</br>");
                        twWriter.Write("<b>Date range: " + " " + report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\'') + " TO " + report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\'') + "</b>");//"+ Fromdate.ToString()+" - " Todate.ToString()+"</b>");
                        twWriter.Write("</br>");
                        twWriter.Write("</br>");

                        twWriter.Write("<div id='Linearcalllist' runat='server' style='border: 1px solid black;overflow: auto;'>");
                        if (report.StoredProcedureName.Contains("spSelectDepartmentalBreakdownreportLevel"))
                        {   // The level is hardcodeed for debug ....
                            //twWriter.Write(HTMLReports.Departmentalreport(tables, GetColumns(report.Columns),"4"));
                            HTMLReports.Departmentalreport(twWriter, tables, GetColumns(report.Columns), report.ListofFilters.Split('=', ',')[93], GetLevels(tables), Listoftotals(report.Totals));
                        }
                        else
                        {
                            if (tables.Tables.Count >= 3)
                            {
                                int i = 0;
                                foreach (DataTable dt in tables.Tables)
                                {
                                    List<string> columns = new List<string>();
                                    foreach (DataColumn ds in dt.Columns)
                                    {
                                        columns.Add(ds.ColumnName);
                                    }
                                    twWriter.Write("<b>" + report.Columns.Split(',')[i] + "</b>");
                                    twWriter.Write(HTMLReports.BindHTMLdata(dt, columns, Listoftotals(report.Totals)));
                                    twWriter.Write("</br>");
                                    i++;
                                }
                            }
                            else
                            {
                                twWriter.Write(HTMLReports.BindHTMLdata(tables.Tables[0], GetColumns(report.Columns), Listoftotals(report.Totals)));
                            }
                        }
                        twWriter.Write("</div>");
                        if (report.GraphBindings != string.Empty)
                        {
                            int repid = 0;
                            foreach (string graph in report.GraphBindings.Split('-'))
                            {
                                GraphicalReport.GenerateGraph("Z:\\inetpub\\wwwroot\\proimages\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", tables.Tables[0], graph, report.ReportName, report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", report.GraphType);
                                twWriter.WriteLine("<div align='center'><img runat='server' width ='900px'  src='http://www.dev.sentelcallmanagerpro.com//images/" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png" + "'/></div>");
                                repid++;
                            }
                        }

                    }
                    else
                    {

                        switch (report.ReportingSection)
                        {
                            #region "KPI"
                            case "Kpi":
                                twWriter.Write("</br>");
                                twWriter.Write("</br>");
                                twWriter.Write("<CENTER><h1>" + report.ReportName + "</h1></CENTER>");
                                twWriter.Write("</br>");
                                twWriter.Write("</br>");
                                twWriter.Write("<b>Schedule Name: " + " " + report.Selectedname + "</b>");
                                twWriter.Write("</br>");
                                twWriter.Write("</br>");
                                twWriter.Write("<b>Date range: " + " " + report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\'') + " TO " + report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\'') + "</b>");//"+ Fromdate.ToString()+" - " Todate.ToString()+"</b>");
                                twWriter.Write("</br>");
                                twWriter.Write("</br>");
                                // Header
                                twWriter.Write(Header.Createheader(tables, report.ReportName, "Kpi"));
                                //Sub headers...( this step is stopped for this release .....)
                                // twWriter.Write(Header.Secondheader(tables, report.ReportName, "Kpi"));

                                //Table body
                                twWriter.Write(HTMLReports.BindHTMLdata(tables.Tables[0], GetColumns(report.Columns), Listoftotals(report.Totals)));
                                // Performance Analysis
                                twWriter.Write(KPI.GenerateKPIReport(tables, report.ReportName, report.UserId));
                                // Graph
                                List<GraphHeaders> Grphe = Datamethods.GetGraphHeadernames(report.GraphHeaders);
                                if (Grphe.Count == 0)
                                {
                                    int repid = 0;
                                    foreach (string graph in report.GraphBindings.Split('-'))
                                    {
                                        GraphicalReport.GenerateGraph("Z:\\inetpub\\wwwroot\\proimages\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", tables.Tables[0], graph, report.ReportName, report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", report.GraphType);
                                        twWriter.WriteLine("<div align='center'><img runat='server' width ='900px'  src='http://www.dev.sentelcallmanagerpro.com//images/" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png" + "'/></div>");
                                        repid++;
                                    }
                                }
                                else
                                {
                                    int repid = 0;
                                    foreach (string graph in report.GraphBindings.Split('-'))
                                    {
                                        foreach (GraphHeaders gr in Grphe)
                                        {
                                            if (graph.Contains(gr.ColumnName))
                                            {
                                                for (int i = 0; i < tables.Tables.Count; i++)
                                                {
                                                    if (tables.Tables[i].Columns.Contains(gr.ColumnName))
                                                    {
                                                        twWriter.WriteLine("<b>" + gr.Headername + "<b>");
                                                        GraphicalReport.GenerateGraph("Z:\\inetpub\\wwwroot\\proimages\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", tables.Tables[i], graph, report.ReportName, report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", report.GraphType);
                                                        twWriter.WriteLine("<div align='center'><img runat='server' width ='900px'  src='http://www.dev.sentelcallmanagerpro.com//images/" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png" + "'/></div>");
                                                    }
                                                }
                                            }
                                            repid++;
                                        }
                                    }
                                }
                                break;
                            #endregion
                            #region "Dashboard"
                            case "Dashboar":
                                twWriter.Write("</br>");
                                twWriter.Write("</br>");
                                twWriter.Write("<CENTER><h1>" + report.ReportName + "</h1></CENTER>");
                                twWriter.Write("</br>");
                                twWriter.Write("</br>");
                                twWriter.Write("<b>Schedule Name: " + " " + report.Selectedname + "</b>");
                                twWriter.Write("</br>");
                                twWriter.Write("</br>");
                                twWriter.Write("<b>Date range: " + " " + report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\'') + " TO " + report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\'') + "</b>");//"+ Fromdate.ToString()+" - " Todate.ToString()+"</b>");
                                twWriter.Write("</br>");
                                twWriter.Write("</br>");
                                // Tables operation.......
                                tables = Tableoperations.Combinetables(tables, report.ReportName, "Dashboar");
                                // Header
                                twWriter.Write(Header.Createheader(tables, report.ReportName, "Dashboar"));
                                //
                                if (report.Columns.Length > 5)
                                {
                                    twWriter.Write(HTMLReports.BindHTMLdata(tables.Tables[0], GetColumns(report.Columns), Listoftotals(report.Totals)));

                                }
                                List<GraphHeaders> Grphead = Datamethods.GetGraphHeadernames(report.GraphHeaders);
                                // Graph
                                if (report.GraphBindings != string.Empty)
                                {
                                    int repid = 0;
                                    foreach (string graph in report.GraphBindings.Split('-'))
                                    {
                                        foreach (GraphHeaders gr in Grphead)
                                        {
                                            if (graph.Contains(gr.ColumnName))
                                            {
                                                for (int i = 0; i < tables.Tables.Count; i++)
                                                {
                                                    if (tables.Tables[i].Columns.Contains(gr.ColumnName))
                                                    {
                                                        string format = report.GraphType;
                                                        if (graph.Split(',').Count() > 2)
                                                        {
                                                            format = "Column";
                                                        }

                                                        twWriter.WriteLine("<b>" + gr.Headername + "<b>");
                                                        GraphicalReport.GenerateGraph("Z:\\inetpub\\wwwroot\\proimages\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", tables.Tables[i], graph, report.ReportName, report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", format);
                                                        twWriter.WriteLine("<div align='center'><img runat='server' width ='900px'  src='http://www.dev.sentelcallmanagerpro.com//images/" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png" + "'/></div>");
                                                    }
                                                }
                                            }
                                            repid++;
                                        }
                                    }
                                }

                                break;
                            #endregion
                            #region "Invoice"
                            case "Invoice":

                                // Generate Header

                                twWriter.Write(Invoice.GenerateUCDInvoiceHeader(tables, report));
                                // Generate Tables
                                twWriter.Write("<div id='Linearcalllist' runat='server' style='border: 1px solid black;overflow: auto;'>");
                                twWriter.Write(Invoice.GenerateUCDInvoiceReport(tables, report, Listoftotals(report.Totals)));
                                twWriter.Write("</div>");
                                break;
                                #endregion
                        }

                    }


                    SendEmail.InsertPortfoliodetails("c:\\temp\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".html", report.EmailAddresses, report.Portfolioreportid, report.ID);


                }
                catch { ReportStatus(report.ID, "ERROR"); }

                finally { twWriter.Close(); }

            }
            else
            {
                // will send service team no data availabe message 
                ReportStatus(report.ID, "No Data");
            }


        }

        // This method will create Word file and inserts into portfolio reports.....
        public static void GenerateWORDReport(Schedule report)
        {

            DataSet tables = ExtecuteReport(report);
            if (tables.Tables[0].Rows.Count >= 1 || report.StoredProcedureName.Contains("spSelectDepartmentalBreakdownreportLevel"))
            {
                TextWriter twWriter = new StreamWriter("c:\\temp\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".doc");

                try
                {
                    twWriter.WriteLine();
                    twWriter.Write("Report Name: " + " " + report.ReportName);
                    twWriter.WriteLine();
                    twWriter.Write("Schedule Name:" + " " + report.Selectedname);
                    twWriter.WriteLine();
                    twWriter.Write("Date range:" + " " + report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\'') + " TO " + report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\''));
                    twWriter.WriteLine();
                    twWriter.WriteLine();
                    if (report.StoredProcedureName.Contains("spSelectDepartmentalBreakdownreportLevel"))
                    {
                        // TextReports.GenerateTXTData(twWriter, " " + '\t', tables.Tables[0], GetColumns(report.Columns), Listoftotals(report.Totals));
                        TextReports.Departmentalreport(twWriter, " " + '\t', tables, GetColumns(report.Columns), report.ListofFilters.Split('=', ',')[93], GetLevels(tables), Listoftotals(report.Totals));
                    }
                    else { TextReports.GenerateTXTData(twWriter, " " + '\t', tables.Tables[0], GetColumns(report.Columns), Listoftotals(report.Totals)); }

                    SendEmail.InsertPortfoliodetails("c:\\temp\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".doc", report.EmailAddresses, report.Portfolioreportid, report.ID);

                }
                catch { ReportStatus(report.ID, "ERROR"); }

                finally { twWriter.Close(); }
            }
            else
            {
                ReportStatus(report.ID, "No Data");
            }


        }
        // This method will create txt report file and inserts into portfolio reports.....
        public static void GenerateTextReport(Schedule report)
        {

            DataSet tables = ExtecuteReport(report);
            if (tables.Tables[0].Rows.Count >= 1 || report.StoredProcedureName.Contains("spSelectDepartmentalBreakdownreportLevel"))
            {
                TextWriter twWriter = new StreamWriter("c:\\temp\\" + report.Selectedname + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".txt");

                try
                {
                    twWriter.WriteLine();
                    twWriter.Write("Report: " + " " + report.ReportName);
                    twWriter.WriteLine();
                    twWriter.Write("Schedule Name:" + " " + report.Selectedname);
                    twWriter.WriteLine();
                    twWriter.Write("Date range:" + " " + report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\'') + " TO " + report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\''));
                    twWriter.WriteLine();
                    twWriter.WriteLine();
                    if (report.StoredProcedureName.Contains("spSelectDepartmentalBreakdownreportLevel"))
                    {
                        // TextReports.GenerateTXTData(twWriter, " " + '\t', tables.Tables[0], GetColumns(report.Columns), Listoftotals(report.Totals));
                        TextReports.Departmentalreport(twWriter, " " + '\t', tables, GetColumns(report.Columns), report.ListofFilters.Split('=', ',')[93], GetLevels(tables), Listoftotals(report.Totals));
                    }
                    else
                    {
                        if (tables.Tables.Count >= 3)
                        {
                            foreach (DataTable dt in tables.Tables)
                            {
                                List<string> columns = new List<string>();
                                foreach (DataColumn dc in dt.Columns)
                                {
                                    columns.Add(dc.ColumnName);
                                }
                                TextReports.GenerateTXTData(twWriter, " " + '\t', dt, columns, Listoftotals(report.Totals));
                                twWriter.WriteLine();
                            }
                        }
                        else
                        {
                            TextReports.GenerateTXTData(twWriter, " " + '\t', tables.Tables[0], GetColumns(report.Columns), Listoftotals(report.Totals));
                        }
                    }
                    SendEmail.InsertPortfoliodetails("c:\\temp\\" + report.Selectedname + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".txt", report.EmailAddresses, report.Portfolioreportid, report.ID);

                }
                catch { ReportStatus(report.ID, "ERROR"); }

                finally { twWriter.Close(); }
            }
            else
            {
                ReportStatus(report.ID, "No Data");
            }


        }

        // This method will create PDF file and inserts into portfolio reports.....
        public static void GeneratePDFReport(Schedule report)
        {    // this check is for departmental reports

            DataSet tables = ExtecuteReport(report);
            if (tables.Tables[0].Rows.Count >= 1 || report.StoredProcedureName.Contains("spSelectDepartmentalBreakdownreportLevel"))
            {
                // create a file with the name
                string file = @"c:\\temp\\ " + report.Selectedname + " - " + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".pdf";
                // creating a pdf file
                iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 5, 5, 10, 10);
                //System.IO.MemoryStream mStream = new System.IO.MemoryStream();            
                PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(file, FileMode.Create));

                try
                {
                    if (Otherreports(report.ReportingSection) == false)
                    {
                        pdfDoc.Open();
                        pdfDoc.Add(new iTextSharp.text.Paragraph("Report Name:" + "-" + report.ReportName, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
                        pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                        pdfDoc.Add(new iTextSharp.text.Paragraph("Schedule Name:" + "-" + report.Selectedname, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 9, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
                        pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                        pdfDoc.Add(new iTextSharp.text.Paragraph("Date range:" + report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\'') + " TO " + report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\''), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 9, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
                        //pdfDoc.Add(new iTextSharp.text.Paragraph(" Date range: " + " -" report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\'') + " TO " + report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\'')));
                        pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                        if (report.StoredProcedureName.Contains("spSelectDepartmentalBreakdownreportLevel"))
                        {
                            // PDFReports.BindPDFdata(pdfDoc, tables.Tables[0], GetColumns(report.Columns), Listoftotals(report.Totals));
                            PDFReports.Departmentalreport(pdfDoc, tables, GetColumns(report.Columns), report.ListofFilters.Split('=', ',')[93], GetLevels(tables), Listoftotals(report.Totals));
                        }
                        else
                        {
                            if (tables.Tables.Count >= 3)
                            {
                                int i = 0;
                                foreach (DataTable dt in tables.Tables)
                                {
                                    List<string> column = new List<string>();
                                    foreach (DataColumn dc in dt.Columns)
                                    {
                                        column.Add(dc.ColumnName);
                                    }
                                    pdfDoc.Add(new iTextSharp.text.Paragraph(report.Columns.Split(',')[i], new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 9, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
                                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                                    PDFReports.BindPDFdata(pdfDoc, dt, column, Listoftotals(report.Totals));
                                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                                    i++;
                                }
                            }
                            else
                            {

                                PDFReports.BindPDFdata(pdfDoc, tables.Tables[0], GetColumns(report.Columns), Listoftotals(report.Totals));
                            }
                        }
                        if (report.GraphBindings != string.Empty)
                        {
                            int repid = 0;
                            foreach (string graph in report.GraphBindings.Split('-'))
                            {
                                GraphicalReport.GenerateGraph("Z:\\inetpub\\wwwroot\\proimages\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", tables.Tables[0], graph, report.ReportName, report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", report.GraphType);
                                iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance("Z:\\inetpub\\wwwroot\\proimages\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png");
                                // scaling into the size
                                jpg.ScaleToFit(450f, 350f);
                                // spacing before image
                                jpg.SpacingBefore = 50f;
                                //spacing  after image
                                jpg.SpacingAfter = 1f;
                                // aligning it ot centre
                                jpg.Alignment = Element.ALIGN_CENTER;
                                // Add to the PDF file
                                pdfDoc.Add(jpg);
                            }
                        }
                    }
                    else
                    {

                        switch (report.ReportingSection)
                        {

                            case "Dashboar":
                                Dashboard.GeneratePDFDashboardData(tables, report, pdfDoc);
                                break;
                            case "Invoice":
                                string costcentrename = tables.Tables[1].Rows[0][0].ToString();
                                string fullname = costcentrename.Split(':')[0] + ":" + report.ListofFilters.Split('=', ',')[9].TrimStart('\'').TrimEnd('\'');
                                List<string> names = new System.Collections.Generic.List<string>();
                                names.Add(fullname);
                                names.Add("Cost Centre:");
                                names.Add("Overall:");
                                names.Add("Destination:");
                                pdfDoc.Open();

                                Invoice.GerateUCDpdfInvoiceHeader(pdfDoc, tables, report);
                                int i = 0;
                                foreach (DataTable dt in tables.Tables)
                                {
                                    List<string> column = new List<string>();
                                    foreach (DataColumn dc in dt.Columns)
                                    {

                                        column.Add(dc.ColumnName);
                                    }
                                    pdfDoc.Add(new iTextSharp.text.Paragraph(names[i], new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
                                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                                    PDFReports.BindPDFdata(pdfDoc, dt, column, Listoftotals(report.Totals));
                                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                                    i++;
                                }

                                pdfDoc.Close();

                                break;
                        }
                    }
                    SendEmail.InsertPortfoliodetails(file, report.EmailAddresses, report.Portfolioreportid, report.ID);

                }
                catch { ReportStatus(report.ID, "ERROR"); }

                finally { pdfDoc.Close(); }
            }
            else
            {
                ReportStatus(report.ID, "No Data");
            }


        }
        // This method will create XML file and inserts into portfolio reports.....
        public static void GenerateXMLReport(Schedule report)
        {


            if (ExtecuteReport(report).Tables[0].Rows.Count >= 1)
            {
                try
                {
                    //  TextWriter twWriter = new StreamWriter("c:\\temp\\" + report.Selectedname + " - " + report.ReportName + DateTime.Now.ToString() + "-" + schReports.ID + ".html");
                }
                catch { }

            }
            else
            {
                ReportStatus(report.ID, "No Data");
            }




        }
        // This method will create CSV report file and inserts into portfolio reports.....
        public static void GenerateCSVReport(Schedule report)
        {


            DataSet tables = ExtecuteReport(report);
            if (tables.Tables[0].Rows.Count >= 1)
            {
                TextWriter twWriter = new StreamWriter("c:\\temp\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".csv", false, Encoding.UTF8);

                try
                {
                    twWriter.WriteLine();
                    twWriter.Write(report.ReportName);
                    twWriter.WriteLine();
                    twWriter.WriteLine();
                    twWriter.Write("Schedule Name:" + " " + report.Selectedname);
                    twWriter.WriteLine();
                    twWriter.WriteLine();
                    twWriter.Write("Date range:" + " " + report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\'') + " TO " + report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\''));
                    twWriter.WriteLine();
                    twWriter.WriteLine();
                    if (report.StoredProcedureName == "spSelectDepartmentalBreakdownreportLevel")
                    {
                        //TextReports.GenerateTXTData(twWriter, ",", tables.Tables[0], GetColumns(report.Columns), Listoftotals(report.Totals));
                        TextReports.Departmentalreport(twWriter, ",", tables, GetColumns(report.Columns), report.ListofFilters.Split('=', ',')[93], GetLevels(tables), Listoftotals(report.Totals));
                    }
                    else
                    {
                        if (tables.Tables.Count >= 3)
                        {
                            int i = 0;
                            foreach (DataTable dt in tables.Tables)
                            {
                                List<string> columns = new List<string>();
                                foreach (DataColumn dc in dt.Columns)
                                {
                                    columns.Add(dc.ColumnName);
                                }
                                twWriter.Write(report.Columns.Split(',')[i]);
                                TextReports.GenerateTXTData(twWriter, ",", dt, columns, Listoftotals(report.Totals));
                                twWriter.WriteLine();
                                i++;
                            }
                        }
                        else
                        {
                            TextReports.GenerateTXTData(twWriter, ",", tables.Tables[0], GetColumns(report.Columns), Listoftotals(report.Totals));
                        }

                    }

                    if (report.GraphBindings != string.Empty)
                    {
                        List<string> pics = new List<string>();
                        int repid = 0;
                        foreach (string graph in report.GraphBindings.Split('-'))
                        {
                            GraphicalReport.GenerateGraph("Z:\\inetpub\\wwwroot\\proimages\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", tables.Tables[0], graph, report.ReportName, report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", report.GraphType);
                            pics.Add("Z:\\inetpub\\wwwroot\\proimages\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png");
                            repid++;
                        }
                        GraphicalReport.CreateGraphsinExcel(pics, "c:\\temp\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".xls", report.ReportName);
                        SendEmail.InsertPortfoliodetails("c:\\temp\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".xls", report.EmailAddresses, report.Portfolioreportid, report.ID);
                    }
                    SendEmail.InsertPortfoliodetails("c:\\temp\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".csv", report.EmailAddresses, report.Portfolioreportid, report.ID);





                }
                catch { ReportStatus(report.ID, "ERROR"); }

                finally { twWriter.Close(); }
            }
            else
            {
                ReportStatus(report.ID, "No Data");
            }




        }
        public static void GenerateExcelReport(Schedule report)
        {
            DataSet tables = ExtecuteReport(report);
            TextWriter twWriter = new StreamWriter("c:\\temp\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".xls");
            try
            {

                if (tables.Tables[0].Rows.Count >= 1)
                {
                    if (Otherreports(report.ReportingSection) == false)
                    {

                    }
                    else
                    {
                        switch (report.ReportingSection)
                        {
                            case "Invoice":
                                string style = @"<style> .textmode { mso-number-format:\@; } </style>";
                                twWriter.Write(style);
                                twWriter.Write(Invoice.GenerateUCDInvoiceHeader(tables, report));
                                // Generate Tables
                                twWriter.Write(Invoice.GenerateUCDInvoiceReport(tables, report, Listoftotals(report.Totals)));
                                break;

                        }

                    }
                }
                else
                {
                    // will send service team no data availabe message 
                    ReportStatus(report.ID, "No Data");
                }
                SendEmail.InsertPortfoliodetails("c:\\temp\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + "-" + report.ID + ".xls", report.EmailAddresses, report.Portfolioreportid, report.ID);
            }
            catch { ReportStatus(report.ID, "ERROR"); }

            finally { twWriter.Close(); }

        }
        private static void Updatereport(Schedule report)
        {
            throw new NotImplementedException();
        }

        public static void ReportStatus(int id, string message)
        {
            tbl_Reportstatus t1 = new tbl_Reportstatus();
            t1.Reportstatus_Status = message;
            t1.Reportstatus_Date = DateTime.Now;
            t1.schedule_id_PK = id;
            Datamethods.DbContext.tbl_Reportstatus.InsertOnSubmit(t1);
            Datamethods.DbContext.SubmitChanges();

        }
        // this method will execute SP and generate dataset.....
        internal static DataSet ExtecuteReport(Schedule report)
        {

            if (report.ListofFilters.Contains("2019-07-01 00"))
            {
                //
                report.ListofFilters = report.ListofFilters.Replace("2019-07-01 00", "2019-06-01 00").Replace("2019-07-31 23", "2019-06-30 23");
                //
            }
            else
            {

            }



            string[] choosennode = report.Chosennodelist.Split('=');
            string param = "'" + choosennode[1] + "'";
            string name = choosennode[0] + "=";
            report.Chosennodelist = name + param;
            SqlConnection oConn = new SqlConnection(ConfigurationManager.AppSettings["TEMConnectionString"]);
            SqlCommand oCmd = new SqlCommand("exec " + report.StoredProcedureName + " " + report.ListofFilters + "," + report.Chosennodelist, oConn);
            // need to change the choosennodelist with ''
            DataSet tables = new DataSet();
            try
            {
                SqlDataAdapter sqlDA = new SqlDataAdapter(oCmd);

                oCmd.CommandTimeout = 30000;
                sqlDA.Fill(tables);

            }
            catch
            {
                ReportStatus(report.ID, "Error");
            }
            finally
            { oConn.Close(); }


            return tables;

        }
        // This method will generate columns......
        private static List<string> GetColumns(string Columns)
        {

            List<string> column = new List<string>();

            foreach (string s in Columns.Split(','))
            {
                column.Add(s.Split('@')[0]);


            }
            return column;

        }
        // This method will get column names which has totals and type of totals....
        public static List<ColumnTotals> Listoftotals(string tot)
        {
            List<ColumnTotals> totals = new List<ColumnTotals>();
            ColumnTotals tt;
            // To check fot totals........
            if (tot != null)
            {
                foreach (string whole in tot.Split(','))
                {
                    tt = new ColumnTotals();
                    tt.Totaltype = whole.Split('@')[1];
                    tt.ColumnName = whole.Split('@')[0];

                    totals.Add(tt);

                }
            }




            return totals;
        }

        private static int GetLevels(DataSet dt)
        {
            if (dt.Tables[6].Rows.Count >= 1)
            {

                return 4;
            }
            else if (dt.Tables[5].Rows.Count >= 1)
            {

                return 3;


            }
            else if (dt.Tables[4].Rows.Count >= 1)
            {

                return 2;

            }
            else
            {

                return 1;
            }


        }

        public static int GetDurationinsec(string duration)
        {
            int i = 0;

            if (duration.Contains(":"))
            {

                i = Convert.ToInt32(duration.Split(':')[0]) * 3600 + Convert.ToInt32(duration.Split(':')[1]) * 60 + Convert.ToInt32(duration.Split(':')[2]);

            }
            return i;
        }
        private static bool Otherreports(string Type)
        {
            switch (Type)
            {
                case "Incoming":
                    return false;

                case "Outgoing":
                    return false;
                case "Combined":
                    return false;
                case "Directory":
                    return false;
                default:
                    return true;

            }


        }
        #region IDisposable Members



        #endregion




        /// <summary>
        /// Handles either a standard logging message or in addition an exception.
        /// Write message to log file, in a specified location
        /// </summary>
        /// <param name="messageText"></param>
        /// <param name="errorException"></param>
        public static void LogMessageToFile(string messageText, Exception errorException = null)
        {
            //Make sure path ends with \\
            //  string path = "C:\\aidan\\ScheduledReportsLogFiles\\";
            string logFilePath = ConfigurationManager.AppSettings["LogFilePath"].ToString();
            string fullMessage = errorException == null ? string.Empty : Environment.NewLine + "ERROR MESSAGE: " + errorException.Message + Environment.NewLine + errorException.StackTrace;
            //  $"{Environment.NewLine}ERROR MESSAGE: {errorException.Message} {Environment.NewLine}STACK TRACE:{Environment.NewLine}{errorException.StackTrace}";

            string logLine = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " " + messageText + fullMessage + Environment.NewLine + "-------------------------------------------------------";
            //$"{Environment.NewLine}{System.DateTime.Now}: {messageText} {fullMessage}  {Environment.NewLine}-------------------------------------------------------";

            using (StreamWriter writer = new StreamWriter(logFilePath + "AutomatedReportsLog.txt", true))
            {
                writer.WriteLine(logLine);
            }
        }



    }
}
