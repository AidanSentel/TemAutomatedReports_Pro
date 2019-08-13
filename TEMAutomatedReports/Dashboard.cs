using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace TEMAutomatedReports
{
    class Dashboard
    {
        public static void GeneratePDFDashboardData(DataSet tables, Schedule report, Document pdfDoc)
        {
            pdfDoc.Open();
            Font link = FontFactory.GetFont("TIMES_ROMAN", 10, Font.BOLD, BaseColor.BLACK);
            Anchor click = new Anchor("Report Name:" + "-" + report.ReportName, link);
            click.Reference = "www.sentelsolutions.com";
            Paragraph p1 = new Paragraph();
            p1.Add(click);
            pdfDoc.Add(p1);
            //pdfDoc.Add(new iTextSharp.text.Paragraph("Report Name:" + "-" + report.ReportName, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
            pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
            pdfDoc.Add(new iTextSharp.text.Paragraph("Schedule Name:" + "-" + report.Selectedname, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 9, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
            pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
            pdfDoc.Add(new iTextSharp.text.Paragraph("Date range:" + report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\'') + " TO " + report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\''), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 9, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
            switch (report.ReportName)
            {
                case "Operator Performance":
                    int repid = 0; List<iTextSharp.text.Image> ImgList = new List<Image>();
                    foreach (DataRow dr in tables.Tables[0].Rows)
                    {
                        GraphicalReport.GenerateCircularGauge("Z:\\inetpub\\wwwroot\\proimages\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png", Convert.ToDouble(dr["Target"]), Convert.ToDouble(dr["_Avg"]), Convert.ToDouble(dr["_Bad"]), Convert.ToDouble(dr["Ring"]), dr[0].ToString(), dr["_Back"].ToString());
                        ImgList.Add(iTextSharp.text.Image.GetInstance("Z:\\inetpub\\wwwroot\\proimages\\" + report.Selectedname + "-" + report.ReportName + DateTime.Now.ToString("ddMMyy") + repid.ToString() + "-" + report.ID + ".png"));
                        repid++;
                    }
                       List<int> numbers = new List<int>();
                       for (int i = 3; i < ImgList.Count; i += 3)
                        {
                          numbers.Add(i);
                        }
                       numbers.Add(numbers.Last()+3);

                       for (int i = ImgList.Count(); i <= numbers.Last(); i++)
                       {

                           ImgList.Add(iTextSharp.text.Image.GetInstance("Z:\\inetpub\\wwwroot\\proimages\\EmptySpace.png"));
                       }
                       

                    var table1 = new PdfPTable(3); //table1
                    table1.HorizontalAlignment = Element.ALIGN_MIDDLE;
                    table1.SpacingBefore = 20;
                    table1.DefaultCell.Border = 0;
                    //table1.WidthPercentage = 20;
                    foreach (Image img in ImgList)
                    {
                        PdfPCell cell = new PdfPCell(img);
                        table1.AddCell(cell);
                    }
                    pdfDoc.Add(table1);
                    PDFReports.BindPDFdataWithKPIIndication(pdfDoc, tables.Tables[0], tables.Tables[0].Columns.Cast<DataColumn>().Where(s => !s.ColumnName.Contains("_") && s.ColumnName != "DrillDown").Select(x => x.ColumnName).ToList(), GenerateReports.Listoftotals(report.Totals), "Ring", "_Back");

                    break;

            }
            pdfDoc.Close();
        }
    }
}
