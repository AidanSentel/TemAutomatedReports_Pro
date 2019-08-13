using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections.Generic;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace TEMAutomatedReports
{
    class Invoice
    {

        internal static string GenerateUCDInvoiceHeader(DataSet ds, Schedule report)
        {
            string dept = report.ListofFilters.Split('=', ',')[9].TrimStart('\'').TrimEnd('\'');
            string costcentrename = ds.Tables[1].Rows[0][0].ToString();
            string cos= costcentrename.Split(':')[0].Replace("Cost Centre", "").Trim();
            StringBuilder HTMLtoRender = new StringBuilder();
            
            try
            {
                HTMLtoRender.Append("<table width='100%'");
                HTMLtoRender.Append("<tr>");
                HTMLtoRender.Append("<td>");
                HTMLtoRender.Append("<table width='600px' style='border:double'>");
                HTMLtoRender.Append("<tr>");
                HTMLtoRender.Append("<td colspan='2'>");
                HTMLtoRender.Append("<h1>" + dept + "</h1>");
                HTMLtoRender.Append("</td>");
                HTMLtoRender.Append("</tr>");
                HTMLtoRender.Append("<tr>");
                HTMLtoRender.Append("<td>");
                HTMLtoRender.Append("Report Type:");
                HTMLtoRender.Append("</td>");
                HTMLtoRender.Append("<td>");
                HTMLtoRender.Append("<b>Customer Account by Fixed Costs</b>");
                HTMLtoRender.Append("</td>");
                HTMLtoRender.Append("</tr>");
                HTMLtoRender.Append("<tr>");
                HTMLtoRender.Append("<td>");
                HTMLtoRender.Append("Cost Centre");
                HTMLtoRender.Append("</td>");
                HTMLtoRender.Append("<td>");
                HTMLtoRender.Append("<b>" + cos + " </b>");
                HTMLtoRender.Append("</td>");
                HTMLtoRender.Append("</tr>");
                HTMLtoRender.Append("<tr>");
                HTMLtoRender.Append("<td>");
                HTMLtoRender.Append("Reporting Period");
                HTMLtoRender.Append("</td>");
                HTMLtoRender.Append("<td>");
                HTMLtoRender.Append("<b>" + string.Format("<b>{0} - {1}</b>", report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\''), report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\'')) + " </b>");
                HTMLtoRender.Append("</td>");
                HTMLtoRender.Append("</tr>");
                HTMLtoRender.Append("</table>");
                HTMLtoRender.Append("</td>");
                HTMLtoRender.Append("<td>");
                HTMLtoRender.Append("<img runat='server' width ='180px' height='120px'  src='http://www.dev.sentelcallmanagerpro.com//images//UCD2.png'/>");
                HTMLtoRender.Append("</td>");
                HTMLtoRender.Append("</tr>");
                HTMLtoRender.Append("<table>");
            }
            catch { }

            return HTMLtoRender.ToString();
        }

        internal static string GenerateUCDInvoiceReport(DataSet ds, Schedule report, List<ColumnTotals> Totals)
        {
            string costcentrename = ds.Tables[1].Rows[0][0].ToString();
            string fullname = costcentrename.Split(':')[0] + ":" + report.ListofFilters.Split('=', ',')[9].TrimStart('\'').TrimEnd('\'');
            List<string> names = new System.Collections.Generic.List<string>();
            names.Add(fullname);
            names.Add("Cost Centre:");
            names.Add("Overall:");
            names.Add("Destination:");
            StringBuilder HTMLtoRender = new StringBuilder();
            try
            {
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    HTMLtoRender.Append("<table>");
                    HTMLtoRender.Append("<tr>");
                    HTMLtoRender.Append("<td ><b>" + names[i].ToString() + "</b></td>");
                    HTMLtoRender.Append("</tr>");
                    HTMLtoRender.Append("</table>");
                    List<string> column = new System.Collections.Generic.List<string>();
                    HTMLtoRender.Append("<table  width='100%' style='border:solid'>");
                    HTMLtoRender.Append("<tr>");
                    foreach (DataColumn dc in ds.Tables[i].Columns)
                    {
                        HTMLtoRender.Append("<td style='border-bottom: 1px solid black;border-right: 1px solid silver;'><b>" + dc.ColumnName + "</b></td>");
                        column.Add(dc.ColumnName);
                    }
                    HTMLtoRender.Append("</tr>");
                    foreach (DataRow dr in ds.Tables[i].Rows)
                    {
                        HTMLtoRender.Append("<tr>");

                        foreach (string s in column)
                        {
                            if (dr[s].ToString().Contains(".") && s.Contains("Cost"))
                            { HTMLtoRender.Append("<td style='border-bottom: 1px solid black;border-right: 1px solid silver;'>" + Math.Round( Convert.ToDecimal( dr[s]),2).ToString() + "</td>"); }
                            else
                            {
                                HTMLtoRender.Append("<td style='border-bottom: 1px solid black;border-right: 1px solid silver;'>" + dr[s].ToString() + "</td>");
                            }
                        }
                        HTMLtoRender.Append("</tr>");
                    }


                    HTMLtoRender.Append("<tr>");
                    foreach (string s in column)
                    {
                        if (Totals.Where(t => t.ColumnName.Contains(s)).Any())
                        {
                            string footer = "";
                            switch (Totals.Where(t => t.ColumnName == s).Select(g => g.Totaltype).SingleOrDefault())
                            {

                                case "Sum":
                                    {
                                        if (s.ToLower().Contains("cost"))
                                        { footer = Math.Round( ds.Tables[i].AsEnumerable().Sum(k => k.Field<decimal>(s)),2).ToString(); }
                                        else
                                        {
                                            footer = ds.Tables[i].AsEnumerable().Sum(y => y.Field<int>(s)).ToString();

                                        }


                                        break;
                                    }
                                case "FormatedSum":
                                    {
                                        var dis = from p in ds.Tables[i].AsEnumerable()
                                                  select new
                                                  {
                                                      avg = p.Field<string>(s)

                                                  };
                                        List<int> lst = new List<int>();
                                        foreach (var grp in dis)
                                        {

                                            lst.Add(((Convert.ToInt32(grp.avg.Substring(0, 2)) * 3600) + (Convert.ToInt32(grp.avg.Substring(3, 2)) * 60) + (Convert.ToInt32(grp.avg.Substring(6, 2)))));

                                        }

                                        TimeSpan t1 = new TimeSpan(0, 0, Convert.ToInt32(lst.Sum()));
                                        footer = t1.ToString();
                                        break;
                                    }
                            }

                            HTMLtoRender.Append("<td style='border-bottom: 1px solid black;border-right: 1px solid silver;'><b> Total:" + footer + "</b></td>");
                        }
                        else { HTMLtoRender.Append("<td style='border-bottom: 1px solid black;border-right: 1px solid silver;'> </td>"); }


                    }
                    HTMLtoRender.Append("</tr>");
                    HTMLtoRender.Append("</table>");
                    HTMLtoRender.Append("</br>");
                }
            }
            catch { }
            return HTMLtoRender.ToString();
        }
        internal static void GerateUCDpdfInvoiceHeader(iTextSharp.text.Document pdfDoc, DataSet ds, Schedule report)
        {
            string dept = report.ListofFilters.Split('=', ',')[9].TrimStart('\'').TrimEnd('\'');
            string costcentrename = ds.Tables[1].Rows[0][0].ToString();
            string cos = costcentrename.Split(':')[0].Replace("Cost Centre", "").Trim();
            PdfPTable pdfTable = new PdfPTable(1);
            PdfPCell cell = new PdfPCell(new Phrase(dept, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
            //cell.Border = iTextSharp.text.Rectangle.NO_BORDER;
            pdfTable.AddCell(cell);
            PdfPCell cell1 = new PdfPCell(new Phrase("Report Type: Customer Account by Fixed Costs", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
           // cell1.Border = iTextSharp.text.Rectangle.NO_BORDER;
            pdfTable.AddCell(cell1);
            PdfPCell cell2 = new PdfPCell(new Phrase("Cost Centre: "+cos, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
            pdfTable.AddCell(cell2);
           // cell2.Border = iTextSharp.text.Rectangle.NO_BORDER;
            PdfPCell cell3 = new PdfPCell(new Phrase("Reporting Period: " + report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\'') + " - " + report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\''), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
            pdfTable.AddCell(cell3);
            
            pdfDoc.Add(pdfTable); 
        }
    }
}
