using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
namespace TEMAutomatedReports
{
    static class HTMLReports
    {

        public static string BindHTMLdata(DataTable dt, List<string> columns , List<ColumnTotals> Totals)
        {
            StringBuilder HTMLtoRender = new StringBuilder("<table class='table table-striped' align ='center' width='100%' style=\"font:15px/20px arial,sans-serif\">", 100000);
            try
            {
                
                // cratating table

                HTMLtoRender.Append("<tr>");
                foreach( string colname in columns){
               HTMLtoRender.AppendFormat("<td style=\"height: 30px;border: 0px none;border-bottom: 1px solid black; border-right: 1px solid silver;padding-left: 3px\"> <b>" + colname + "</b></td>", "100px");
                   
                  
	
	

                }
                HTMLtoRender.Append("</tr>");
                foreach (DataRow row in dt.Rows)
                {
                    HTMLtoRender.Append("<tr>");
                    foreach (string col in columns)
                    {



                        HTMLtoRender.AppendFormat("<td style=\"height: 25px; border-bottom: 1px solid silver;border-right: 1px solid silver;padding-left: 3px;\">", "150px");
                        HTMLtoRender.Append(row[col].ToString());
                        HTMLtoRender.Append("</td>");
                    }
                    HTMLtoRender.Append("</tr>");
                }
                // need to write code for Total here
                
               
               
                HTMLtoRender.Append("<tr>");

            }
            catch { }
                foreach (string col in columns)
                {
                    
                    double footer = 0;
                    try
                    {
                        HTMLtoRender.AppendFormat("<td>", "100px");
                        if (Totals.Where(t => t.ColumnName.Contains(col)).Any())
                        {

                            switch (Totals.Where(t => t.ColumnName.Contains(col)).Select(g => g.Totaltype).SingleOrDefault())
                            {
                                case "Count":
                                    {
                                        footer = dt.AsEnumerable().Count();
                                        HTMLtoRender.Append("Count" + ":" + footer.ToString());
                                        break;
                                    }
                                case "Sum":
                                    {
                                        if (col.ToLower().Contains("cost"))
                                        { HTMLtoRender.Append("Total" + ":" + dt.AsEnumerable().Sum(s => s.Field<decimal>(col)).ToString()); }
                                        else
                                        {
                                            footer = dt.AsEnumerable().Sum(s => s.Field<int>(col));
                                            HTMLtoRender.Append("Total" + ":" + footer.ToString());
                                        }


                                        break;
                                    }
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

                                            lst.Add(((Convert.ToInt32(grp.avg.Substring(0, 2)) * 3600) + (Convert.ToInt32(grp.avg.Substring(3, 2)) * 60) + (Convert.ToInt32(grp.avg.Substring(6, 2)))));

                                        }

                                        TimeSpan t1 = new TimeSpan(0, 0, Convert.ToInt32(lst.Sum()));
                                        HTMLtoRender.Append("Total" + ":" + t1.ToString());
                                        break;
                                    }
                                case "Avg":
                                    {
                                        footer = dt.AsEnumerable().Average(s => s.Field<int>(col));
                                        HTMLtoRender.Append("Avg" + ":" + Math.Round( footer,2).ToString());
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
                                        HTMLtoRender.Append("Avg" + ":" + t1.ToString());
                                        break;
                                    }
                            }
                        }
                    }
                    catch { }
                    }
                    HTMLtoRender.Append("</td>");
                
               

                HTMLtoRender.Append("</tr>");
                HTMLtoRender.Append("</table>");


            return HTMLtoRender.ToString(); 
        
        
        
        }


        public static void Departmentalreport(TextWriter twWriter, DataSet dt, List<string> columns, string level, int direclevel, List<ColumnTotals> Totals)
        {

            if (level == "4")
            {
                twWriter.Write(WriteHTMlTable(dt.Tables[2], 0, Totals));
                // int tablevalue = GetLevels(dt);
                if (direclevel != 1)
                {
                    twWriter.Write(WriteHTMlTable(dt.Tables[direclevel + 2], 1, Totals));
                }
                else
                {
                    twWriter.Write(WriteHTMlTable(dt.Tables[4], 1, Totals));
                    twWriter.Write(WriteHTMlextTable(dt.Tables[0], dt.Tables[3], Totals));
                }
            }
            else
            {
                if (direclevel == 1)
                {
                    twWriter.Write(WriteHTMlTable(dt.Tables[3], 0, Totals));
                }
                else if (direclevel == 2)
                {
                    twWriter.Write(WriteHTMlTable(dt.Tables[2], 0, Totals));
                    twWriter.Write(WriteHTMlTable(dt.Tables[4], 0, Totals));
                    twWriter.Write(WriteHTMlextTable(dt.Tables[0], dt.Tables[3], Totals));
                }
                else if (direclevel == 3)
                {
                    twWriter.Write(WriteHTMlTable(dt.Tables[2], 0, Totals));
                    twWriter.Write(WriteHTMlTable(dt.Tables[5], 1, Totals));
                    twWriter.Write(WriteHTMlTable(dt.Tables[4], 1, Totals));
                    twWriter.Write(WriteHTMlextTable(dt.Tables[0], dt.Tables[3], Totals));
                }
                else if (direclevel == 4)
                {
                    twWriter.Write(WriteHTMlTable(dt.Tables[2], 0, Totals));
                    twWriter.Write(WriteHTMlTable(dt.Tables[6], 1, Totals));
                    twWriter.Write(WriteHTMlTable(dt.Tables[5], 1, Totals));
                    twWriter.Write(WriteHTMlTable(dt.Tables[4], 1, Totals));
                    twWriter.Write(WriteHTMlextTable(dt.Tables[0], dt.Tables[3], Totals));
                }
            }






        }
        private static string WriteHTMlTable(DataTable thisTable, int value, List<ColumnTotals> Totals)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder("<table class='table table-striped' align ='center' width='100%' style=\"font:15px/20px arial,sans-serif\">", 100000);
            try
            {
                if (value == 0)
                {
                    thisTable.Columns.RemoveAt(0);

                }

                sb.Append("<TR>");

                //first append the column names.
                foreach (DataColumn column in thisTable.Columns)
                {
                    if (column.ColumnName != "Totalnonform" && column.ColumnName != "indurnonformated" && column.ColumnName != "outdurnonformated")
                    {
                        sb.Append("<TD><B>");
                        sb.Append(column.ColumnName);
                        sb.Append("</B></TD>");
                    }
                }

                sb.Append("</TR>");

                // next, the column values.
                foreach (DataRow row in thisTable.Rows)
                {
                    sb.Append("<TR>");

                    foreach (DataColumn column in thisTable.Columns)
                    {
                        if (column.ColumnName != "Totalnonform" && column.ColumnName != "indurnonformated" && column.ColumnName != "outdurnonformated")
                        {
                            sb.Append("<TD>");
                            if (row[column].ToString().Trim().Length > 0)
                                sb.Append(row[column]);
                            else
                                sb.Append(" ");
                            sb.Append("</TD>");
                        }
                    }

                    sb.Append("</TR>");
                }


                sb.Append("<TR>");
                try
                {

                    foreach (DataColumn colddd in thisTable.Columns)
                    {
                        string col = colddd.ColumnName;
                        double footer = 0;
                        if (colddd.ColumnName != "Totalnonform" && colddd.ColumnName != "indurnonformated" && colddd.ColumnName != "outdurnonformated")
                        {
                            sb.AppendFormat("<td>", "100px");
                            if (Totals.Where(t => t.ColumnName.Contains(col)).Any())
                            {

                                switch (Totals.Where(t => t.ColumnName.Contains(col)).Select(g => g.Totaltype).SingleOrDefault())
                                {
                                    case "Count":
                                        {
                                            footer = thisTable.AsEnumerable().Count();
                                            sb.Append("Count" + ":" + footer.ToString());
                                            break;
                                        }
                                    case "Sum":
                                        {
                                            if (col.ToLower().Contains("cost"))
                                            { sb.Append("Total" + ":" + thisTable.AsEnumerable().Sum(s => s.Field<decimal>(col)).ToString()); }
                                            else
                                            {
                                                footer = thisTable.AsEnumerable().Sum(s => s.Field<int>(col));
                                                sb.Append("Total" + ":" + footer.ToString());
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
                                            int dis = thisTable.AsEnumerable().Sum(s => s.Field<int>(col));



                                            TimeSpan t1 = new TimeSpan(0, 0, dis);
                                            sb.Append("Total" + ":" + t1.ToString());
                                            break;
                                        }
                                    case "Avg":
                                        {
                                            footer = Math.Round( thisTable.AsEnumerable().Average(s => s.Field<int>(col)),2);
                                            sb.Append("Avg" + ":" + footer.ToString());
                                            break;
                                        }
                                    case "FormatedAvg":
                                        {
                                            var dis = from p in thisTable.AsEnumerable()
                                                      select new
                                                      {
                                                          avg = p.Field<string>(col)

                                                      };
                                            List<int> lst = new List<int>();
                                            foreach (var grp in dis)
                                            {

                                                lst.Add(((Convert.ToInt32(grp.avg.Substring(0, 2)) * 3600) + (Convert.ToInt32(grp.avg.Substring(3, 2)) * 60) + (Convert.ToInt32(grp.avg.Substring(6, 2)))));

                                            }

                                            //TimeSpan t1 = new TimeSpan(0, 0, Convert.ToInt32(Math.Round(lst.Average())));
                                            // sb.Append("Avg" + ":" + t1.ToString());
                                            break;
                                        }
                                }
                            }
                            sb.Append("</td>");
                        }
                    }

                    sb.Append("</TR>");
                }
                catch { }
                sb.Append("</table>");
            }
            catch { }
            return sb.ToString();

        }
        private static StringBuilder WriteHTMlextTable(DataTable des, DataTable dt, List<ColumnTotals> Totals)
        {




            var depts = dt.AsEnumerable().Select(s => s.Field<string>("Cost Centre")).Distinct();
            //StringBuilder HTMLtoRender = new StringBuilder("<table class='table table-striped' align ='center' width='100%' style=\"font:15px/20px arial,sans-serif\">", 100000);
            StringBuilder HTMLtoRender = new StringBuilder();

            try
            {
                foreach (var v in depts)
                {
                    HTMLtoRender.Append("<table class='table table-striped' align ='center' width='100%' style=\"font:15px/20px arial,sans-serif\">");
                    HTMLtoRender.Append("<tr>");
                    HTMLtoRender.AppendFormat(" <td style=\"height: 30px;border: 0px none;border-bottom: 1px solid black; border-right: 1px solid silver;padding-left: 3px\"> <b>");

                    HTMLtoRender.Append(v.ToString());
                    HTMLtoRender.Append("<br />");
                    HTMLtoRender.Append("<br />");
                    HTMLtoRender.AppendFormat("</b></td>");
                    HTMLtoRender.Append("</tr>");
                    HTMLtoRender.Append("<br />");
                    HTMLtoRender.Append("<br />");
                    HTMLtoRender.Append("</table>");
                    HTMLtoRender.Append("<table class='table table-striped' align ='center' width='100%' style=\"font:15px/20px arial,sans-serif\">");
                    HTMLtoRender.Append("<tr>");
                    foreach (DataColumn column in des.Columns)
                    {
                        if (column.ColumnName != "Department" && column.ColumnName != "Totalnonform")
                        {
                            HTMLtoRender.AppendFormat("<td ><B>" + column.ColumnName + "</B></td>");
                        }

                    }
                    HTMLtoRender.Append("</tr>");

                    var dest = from db in des.AsEnumerable()
                               where db.Field<string>("Department") == v.ToString()
                               select new
                               {
                                   Destinationname = db.Field<object>("Destination name"),
                                   Totalcalls = db.Field<object>("Total calls"),
                                   Totalduration = db.Field<object>("Total duration"),
                                   Totalnonform = db.Field<object>("Totalnonform"),
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
                    #region problem with totals code...........
                    HTMLtoRender.Append("<TR>");
                    try
                    {
                        foreach (DataColumn colddd in des.Columns)
                        {

                            string col = colddd.ColumnName;
                            if (col != "Department" && col != "Totalnonform")
                            {
                                string footer = "";
                                HTMLtoRender.AppendFormat("<td>", "100px");
                                if (Totals.Where(t => t.ColumnName.Contains(col)).Any())
                                {

                                    switch (Totals.Where(t => t.ColumnName.Contains(col)).Select(g => g.Totaltype).SingleOrDefault())
                                    {
                                        case "Count":
                                            {
                                                footer = Convert.ToString(dest.AsEnumerable().Count());
                                                HTMLtoRender.Append("Count" + ":" + footer.ToString());
                                                break;
                                            }
                                        case "Sum":
                                            {
                                                if (col.ToLower().Contains("cost"))
                                                { HTMLtoRender.Append("Total" + ":" + Math.Round( dest.AsEnumerable().Sum(s => Convert.ToDouble(s.Cost)),3).ToString()); }
                                                else
                                                {
                                                    footer = Convert.ToString(dest.AsEnumerable().Sum(s => Convert.ToInt32(s.Totalcalls)));
                                                    HTMLtoRender.Append("Total" + ":" + footer);
                                                }


                                                break;
                                            }
                                        case "FormatedSum":
                                            {
                                                int form = dest.AsEnumerable().Sum(s => Convert.ToInt32(s.Totalnonform));
                                                TimeSpan t1 = new TimeSpan(0, 0, form);
                                                HTMLtoRender.Append("Total" + ":" + t1.ToString());
                                                break;
                                            }


                                    }

                                } HTMLtoRender.Append("</td>");
                            }
                        }
                    }
                    catch { }
                    HTMLtoRender.Append("</TR>");


                    #endregion

                    HTMLtoRender.Append("</BR>");
                    HTMLtoRender.Append("</table>");
                    HTMLtoRender.Append("<table class='table table-striped' align ='center' width='100%' style=\"font:15px/20px arial,sans-serif\">");
                    HTMLtoRender.Append("<tr >");
                    foreach (DataColumn column in dt.Columns)
                    {
                        if (column.ColumnName != "Cost Centre" && column.ColumnName != "indurnonformated" && column.ColumnName != "outdurnonformated")
                        {
                            HTMLtoRender.AppendFormat("<td ><B>" + column.ColumnName + "</B></td>");
                        }


                    }
                    HTMLtoRender.Append("</tr>");


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
                        HTMLtoRender.Append("<tr>");
                        HTMLtoRender.AppendFormat("<td>", "150px");
                        HTMLtoRender.Append(all.Name.ToString().Replace(" ", string.Empty));
                        HTMLtoRender.Append("</td>");
                        HTMLtoRender.AppendFormat("<td>", "150px");
                        HTMLtoRender.Append(all.Extension.ToString().Replace(" ", string.Empty));
                        HTMLtoRender.Append("</td>");
                        HTMLtoRender.AppendFormat("<td >", "150px");
                        HTMLtoRender.Append(all.OutgoingCalls.ToString().Replace(" ", string.Empty));
                        HTMLtoRender.Append("</td>");
                        HTMLtoRender.AppendFormat("<td >", "150px");
                        HTMLtoRender.Append(all.OutgoingDuration.ToString().Replace(" ", string.Empty));
                        HTMLtoRender.Append("</td>");
                        HTMLtoRender.AppendFormat("<td >", "150px");
                        HTMLtoRender.Append(all.RingResponse.ToString().Replace(" ", string.Empty));
                        HTMLtoRender.Append("</td>");
                        HTMLtoRender.AppendFormat("<td>", "150px");
                        HTMLtoRender.Append(all.AbandonedCalls.ToString().Replace(" ", string.Empty));
                        HTMLtoRender.Append("</td>");
                        HTMLtoRender.AppendFormat("<td>", "150px");
                        HTMLtoRender.Append(all.IncomingCalls.ToString().Replace(" ", string.Empty));
                        HTMLtoRender.Append("</td>");
                        HTMLtoRender.AppendFormat("<td >", "150px");
                        HTMLtoRender.Append(all.IncomingDuration.ToString().Replace(" ", string.Empty));
                        HTMLtoRender.Append("</td>");
                        HTMLtoRender.AppendFormat("<td>", "150px");
                        HTMLtoRender.Append(all.Cost.ToString().Replace(" ", string.Empty));
                        HTMLtoRender.Append("</td>");
                        HTMLtoRender.Append("</tr>");
                    }
                    # region "not working need more attention..........."
                    HTMLtoRender.Append("<TR>");
                    try
                    {
                        foreach (DataColumn colddd in dt.Columns)
                        {

                            string col = colddd.ColumnName;
                            if (col != "Cost Centre" && col != "indurnonformated" && col != "outdurnonformated")
                            {
                                string footer = "0";
                                HTMLtoRender.AppendFormat("<td>", "100px");
                                if (Totals.Where(t => t.ColumnName.Contains(col)).Any())
                                {

                                    switch (Totals.Where(t => t.ColumnName.Contains(col)).Select(g => g.Totaltype).SingleOrDefault())
                                    {
                                        case "Count":
                                            {
                                                footer = Convert.ToString(records.AsEnumerable().Count());
                                                HTMLtoRender.Append("Count" + ":" + footer);
                                                break;
                                            }
                                        case "Sum":
                                            {
                                                if (col.ToLower().Contains("cost"))
                                                { HTMLtoRender.Append("Total" + ":" + Convert.ToString(Math.Round( records.AsEnumerable().Sum(s => Convert.ToDouble(s.Cost)),3))); }
                                                else if (col.Contains("Incoming Calls"))
                                                {
                                                    footer = Convert.ToString(records.AsEnumerable().Sum(s => Convert.ToInt32(s.IncomingCalls)));
                                                    HTMLtoRender.Append("Total" + ":" + footer.ToString());
                                                }
                                                else if (col.Contains("Outgoing Calls"))
                                                {
                                                    footer = Convert.ToString(records.AsEnumerable().Sum(s => Convert.ToInt32(s.OutgoingCalls)));
                                                    HTMLtoRender.Append("Total" + ":" + footer.ToString());
                                                }

                                                break;
                                            }
                                        case "FormatedSum":
                                            {
                                                int dur = 0; TimeSpan t1;
                                                if (col.Contains("Incoming Duration"))
                                                {
                                                    dur = records.Sum(s => Convert.ToInt32(s.Innonformated));
                                                }
                                                else
                                                {
                                                    dur = records.Sum(s => Convert.ToInt32(s.outnonformated));
                                                }
                                                t1 = new TimeSpan(0, 0, dur);
                                                HTMLtoRender.Append("Total" + ":" + t1.ToString());
                                                break;
                                            }
                                        case "Avg":
                                            {
                                                footer = Convert.ToString(Math.Round(records.AsEnumerable().Average(s => Convert.ToInt32(s.RingResponse)),2));
                                                HTMLtoRender.Append("Avg" + ":" + footer.ToString());
                                                break;
                                            }

                                    }
                                }
                                HTMLtoRender.Append("</td>");
                            }
                        }

                        HTMLtoRender.Append("</TR>");
                    }
                    catch
                    { }
                    #endregion
                    HTMLtoRender.Append("</table>");
                }
            }
            catch { }

            return HTMLtoRender;
        }



    }
}
