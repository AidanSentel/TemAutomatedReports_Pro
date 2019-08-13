using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace TEMAutomatedReports
{
    class PDFReports
    {

        public static void BindPDFdata(iTextSharp.text.Document pdfDoc, DataTable dt, List<string> columns, List<ColumnTotals> Totals)
        {

            PdfPTable pdfTable = new PdfPTable(columns.Count);
            pdfTable.TotalWidth = 100;
            //creating header columns
            foreach (string colu in columns)
            {

                PdfPCell cell = new PdfPCell(new Phrase(colu, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 6, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
                   
                    pdfTable.AddCell(cell);
                
            }
            //creating rows 
            foreach (DataRow row in dt.Rows)
            {
                foreach (string col in columns)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(row[col].ToString().Replace("<b>", string.Empty).Replace("</b>", string.Empty), new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 6, 0)));
                    pdfTable.AddCell(cell);
                }
            }

            // This code is for totals................
            // This code is for totals................
            foreach (string col in columns)
            {
                // double footer = 0;
                if (Totals.Where(t => t.ColumnName == col).Any())
                {
                    string s = Getvalue(Totals.Where(t => t.ColumnName == col).Select(g => g.Totaltype).SingleOrDefault(), dt, col);
                    PdfPCell cell = new PdfPCell(new Phrase(s, new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8, 0)));
                    pdfTable.AddCell(cell);
                }
                else
                {
                    PdfPCell cell = new PdfPCell(new Phrase(" ", new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8, 0)));
                    pdfTable.AddCell(cell);
                }
            }
            pdfDoc.Add(pdfTable);         
        
        }
        private static string  Getvalue(string total, DataTable dt , string col )
        {
            string footer = "0";
            try
            {
             
                switch (total)
                {
                    case "Count":
                        {
                            footer = "Count:" + dt.Rows.Count.ToString();
                            break;
                        }
                    case "Sum":

                        if (col.ToLower().Contains("cost"))
                        {
                            footer = "Total:" + dt.AsEnumerable().Sum(s => s.Field<decimal>(col)).ToString();
                        }
                        else
                        {
                            footer = "Total:" + dt.AsEnumerable().Sum(s => s.Field<int>(col)).ToString();

                        }


                        break;



                    case "FormatedSum":
                        {
                            var dis = from p in dt.AsEnumerable()
                                      select new
                                      {
                                          avg = p.Field<string>(col)

                                      };
                            List<int> lst = new List<int>();
                            foreach (var grp in dis)
                            {
                                if (grp.avg != "0")
                                {
                                    lst.Add(((Convert.ToInt32(grp.avg.Substring(0, 2)) * 3600) + (Convert.ToInt32(grp.avg.Substring(3, 2)) * 60) + (Convert.ToInt32(grp.avg.Substring(6, 2)))));
                                }
                            }

                            TimeSpan t1 = new TimeSpan(0, 0, Convert.ToInt32(lst.Sum()));
                            footer = "Total:" + t1.ToString();

                            break;
                        }
                    case "Avg":
                        {
                            footer = "Avg:" + dt.AsEnumerable().Average(s => s.Field<int>(col)).ToString();

                            break;
                        }
                    case "FormatedAvg":
                        {
                            var dis = from p in dt.AsEnumerable()
                                      select new
                                      {
                                          avg = p.Field<string>(col)

                                      };
                            List<int> lst = new List<int>();
                            foreach (var grp in dis)
                            {
                                if (grp.avg != "0")
                                {
                                    lst.Add(((Convert.ToInt32(grp.avg.Substring(0, 2)) * 3600) + (Convert.ToInt32(grp.avg.Substring(3, 2)) * 60) + (Convert.ToInt32(grp.avg.Substring(6, 2)))));
                                }
                            }

                            TimeSpan t1 = new TimeSpan(0, 0, Convert.ToInt32(Math.Round(lst.Average())));
                            footer = "Avg:" + t1.ToString();
                            break;
                        }
                }

            }
            catch { }

            return footer;
        }
        private static string Getvalue(string total, DataTable thisTable, string col, bool val)
        {
            string footer = "0";
            try
            {


                switch (total)
                {
                    case "Count":
                        {
                            footer = "Count:" + thisTable.Rows.Count.ToString();

                            break;
                        }
                    case "Sum":
                        {
                            if (col.ToLower().Contains("cost"))
                            { footer = "Total:" + thisTable.AsEnumerable().Sum(s => Convert.ToDecimal(s.Field<object>(col))).ToString(); }
                            else
                            {
                                var v = thisTable.AsEnumerable().Select(s => s.Field<object>(col));
                                double g = thisTable.AsEnumerable().Sum(s => Convert.ToDouble(s.Field<object>(col)));

                                footer = "Total:" + thisTable.AsEnumerable().Sum(s => Convert.ToDouble(s.Field<object>(col))).ToString();

                            }


                            break;
                        }
                    case "FormatedSum":
                        {
                            switch (col)
                            {
                                case "Incoming Duration":
                                    col = "indurnonformated";
                                    break;
                                case "Outgoing Duration":
                                    col = "outdurnonformated";
                                    break;
                                case "Total Duration":
                                    col = "Totalnonform";
                                    break;

                            }
                            int dis = thisTable.AsEnumerable().Sum(s => Convert.ToInt32(s.Field<object>(col)));



                            TimeSpan t1 = new TimeSpan(0, 0, dis);
                            footer = "Total:" + t1.ToString();
                            break;
                        }
                    case "Avg":
                        {
                            footer = "Avg:" + Math.Round( thisTable.AsEnumerable().Average(s => Convert.ToInt32(s.Field<object>(col))),2).ToString();

                            break;
                        }

                }
            }






            catch { }
            return footer;

        }
        public static void GetPDfTable(iTextSharp.text.Document pdfDoc, DataTable dataTable, int level, bool first, List<ColumnTotals> Totals)
        {

            int cols = dataTable.Columns.Count;

            int rows = dataTable.Rows.Count;
            if (first == true)
            {
                cols = cols - 2;
            }
            else
            { cols = cols - 2; }
            PdfPTable pdfTable = new PdfPTable(cols);
            pdfTable.TotalWidth = 100;
            List<string> column = new List<string>();
            //creating header columns
            foreach (DataColumn colu in dataTable.Columns)
            {
                if (colu.ColumnName != "Department" && colu.ColumnName != "Totalnonform" && colu.ColumnName != "indurnonformated" && colu.ColumnName != "outdurnonformated")
                {
                    PdfPCell cell = new PdfPCell(new Phrase(colu.ColumnName, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));
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

            foreach (string col in column)
            {
                //double footer = 0;
                if (Totals.Where(t => t.ColumnName.Contains(col)).Any())
                {
                    //string s = Getvalue(Totals.Where(t => t.ColumnName.Contains(col)).Select(g => g.Totaltype).SingleOrDefault(), dataTable, col);

                    string s = Getvalue(Totals.Where(t => t.ColumnName.Contains(col)).Select(g => g.Totaltype).SingleOrDefault(), dataTable, col, true);
                    PdfPCell cell = new PdfPCell(new Phrase(s, new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8, 0)));
                    pdfTable.AddCell(cell);
                }
                else
                {
                    PdfPCell cell = new PdfPCell(new Phrase(" ", new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8, 0)));
                    pdfTable.AddCell(cell);
                }
            }
            pdfDoc.Add(pdfTable);


        }
        public static void GetEXTPDFTable(iTextSharp.text.Document pdfDoc, DataTable des, DataTable dt, List<ColumnTotals> Totals)
        {
            try
            {
                DataTable table1 = new DataTable();
                DataTable table2 = new DataTable();

                table1.Columns.Clear();
                table2.Columns.Clear();
                var depts = dt.AsEnumerable().Select(s => s.Field<string>("Cost Centre")).Distinct();
                foreach (DataColumn colu in des.Columns)
                {

                    // if (colu.ColumnName != "Department")
                    // {
                    table1.Columns.Add(colu.ColumnName);
                    //}
                }
                foreach (DataColumn colu in dt.Columns)
                {
                    if (colu.ColumnName != "Cost Centre")
                    {
                        table2.Columns.Add(colu.ColumnName);
                    }
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
                                   Department = db.Field<object>("Department"),
                                   Destinationname = db.Field<object>("Destination Name"),
                                   Totalcalls = db.Field<object>("Total Calls"),
                                   Totalduration = db.Field<object>("Total Duration"),
                                   Totalnonform = db.Field<object>("Totalnonform"),
                                   Cost = db.Field<object>("Cost")

                               };
                    foreach (var destinations in dest)
                    {
                        table1.Rows.Add(destinations.Department.ToString(), destinations.Destinationname.ToString(), destinations.Totalcalls.ToString(), destinations.Totalduration.ToString(), destinations.Totalnonform.ToString(), destinations.Cost.ToString());


                    }
                    GetPDfTable(pdfDoc, table1, 0, false, Totals);

                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    var records = from db in dt.AsEnumerable()
                                  where db.Field<string>("Cost Centre") == v.ToString()
                                  select new
                                  {
                                      Name = db.Field<object>("Name"),
                                      Extension = db.Field<object>("Extension"),
                                      OutgoingCalls = db.Field<object>("Outgoing Calls"),
                                      OutgoingDuration = db.Field<object>("Outgoing Duration"),
                                      outnonformated = db.Field<object>("outdurnonformated"),
                                      RingResponse = db.Field<object>("Ring Response"),
                                      AbandonedCalls = db.Field<object>("Abandoned Calls"),
                                      IncomingCalls = db.Field<object>("Incoming Calls"),
                                      IncomingDuration = db.Field<object>("Incoming Duration"),
                                      Innonformated = db.Field<object>("indurnonformated"),
                                      Cost = db.Field<object>("Cost")

                                  };
                    foreach (var all in records)
                    {
                        table2.Rows.Add(all.Name.ToString(), all.Extension.ToString(), all.OutgoingCalls.ToString(), Convert.ToString(all.OutgoingDuration), Convert.ToString(all.outnonformated), all.RingResponse.ToString(), all.AbandonedCalls.ToString(), all.IncomingCalls.ToString(), Convert.ToString(all.IncomingDuration), Convert.ToString(all.Innonformated), Convert.ToString(all.Cost));

                    }

                    GetPDfTable(pdfDoc, table2, 0, false, Totals);

                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                }






            }
            catch { }




        }
        public static void Departmentalreport( iTextSharp.text.Document pdfDoc, DataSet dt, List<string> columns, string level, int direclevel, List<ColumnTotals> Totals)
        {
         if (level == "4")
            {
                GetPDfTable(pdfDoc, dt.Tables[2], 0, true, Totals);
                pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                if (direclevel != 1)
                {
                    GetPDfTable(pdfDoc, dt.Tables[direclevel + 2], 1, false, Totals);
                }
                else { GetEXTPDFTable(pdfDoc, dt.Tables[0], dt.Tables[3], Totals); }
            }
            else
            {
                if (direclevel == 1)
                {
                    GetPDfTable(pdfDoc, dt.Tables[3], 0, true, Totals);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                }
                else if (direclevel == 2)
                {
                    GetPDfTable(pdfDoc, dt.Tables[2], 0, true, Totals);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[4], 0, false, Totals);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetEXTPDFTable(pdfDoc, dt.Tables[0], dt.Tables[3], Totals);
                }
                else if (direclevel == 3)
                {
                    GetPDfTable(pdfDoc, dt.Tables[2], 0, true, Totals);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[5], 1, false, Totals);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[4], 1, false, Totals);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetEXTPDFTable(pdfDoc, dt.Tables[0], dt.Tables[3], Totals);
                }
                else if (direclevel == 4)
                {
                    GetPDfTable(pdfDoc, dt.Tables[2], 0, true, Totals);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[6], 1, false, Totals);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[5], 1, false, Totals);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetPDfTable(pdfDoc, dt.Tables[4], 1, false, Totals);
                    pdfDoc.Add(new iTextSharp.text.Paragraph(" "));
                    GetEXTPDFTable(pdfDoc, dt.Tables[0], dt.Tables[3], Totals);
                }
            }

        }
        public static void BindPDFdataWithKPIIndication(iTextSharp.text.Document pdfDoc, DataTable dt, List<string> columns, List<ColumnTotals> Totals, string targetcolumn, string colourcode)
        {
            PdfPTable pdfTable = new PdfPTable(columns.Count);
            pdfTable.TotalWidth = 100;
            //creating header columns
            foreach (string colu in columns)
            {

                PdfPCell cell = new PdfPCell(new Phrase(colu, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 6, iTextSharp.text.Font.BOLD, BaseColor.BLACK)));

                pdfTable.AddCell(cell);

            }
            //creating rows 
            foreach (DataRow row in dt.Rows)
            {
                foreach (string col in columns)
                {
                    if (col == targetcolumn)
                    {
                        PdfPCell cell = new PdfPCell(new Phrase(row[col].ToString().Replace("<b>", string.Empty).Replace("</b>", string.Empty), new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 6, 0)));
                        BaseColor myColor = new BaseColor(255, 255, 255);
                        switch (row[colourcode].ToString())
                        {

                            case "Table_Rows_Green":
                                myColor = new BaseColor(141, 203, 42);

                                break;
                            case "Table_Rows_orange":
                                myColor = new BaseColor(255, 122, 0);

                                break;
                            case "Table_Rows_Red":
                                myColor = new BaseColor(194, 0, 0);

                                break;
                        }
                        cell.BackgroundColor = myColor;
                        pdfTable.AddCell(cell);
                    }
                    else
                    {
                        PdfPCell cell = new PdfPCell(new Phrase(row[col].ToString().Replace("<b>", string.Empty).Replace("</b>", string.Empty), new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 6, 0)));
                        pdfTable.AddCell(cell);
                    }
                }
            }

            // This code is for totals................
            foreach (string col in columns)
            {
                // double footer = 0;
                if (Totals.Where(t => t.ColumnName == col).Any())
                {
                    string s = Getvalue(Totals.Where(t => t.ColumnName == col).Select(g => g.Totaltype).SingleOrDefault(), dt, col);
                    PdfPCell cell = new PdfPCell(new Phrase(s, new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8, 0)));
                    pdfTable.AddCell(cell);
                }
                else
                {
                    PdfPCell cell = new PdfPCell(new Phrase(" ", new iTextSharp.text.Font(iTextSharp.text.Font.NORMAL, 8, 0)));
                    pdfTable.AddCell(cell);
                }
            }
            pdfDoc.Add(pdfTable);

        }
    }
}
