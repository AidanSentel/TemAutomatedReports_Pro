using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;


namespace TEMAutomatedReports
{
    class TextReports
    {
        public static void GenerateTXTData(TextWriter twWriter, string seperator, DataTable dt, List<string> column, List<ColumnTotals> Totals)
        {
            try
            {

                //creating table headers
                foreach (string colu in column)
                {

                    twWriter.Write(colu);
                    twWriter.Write(seperator);

                }
                twWriter.WriteLine();
                //Creating columns
                foreach (DataRow row in dt.Rows)
                {
                    foreach (string col in column)
                    {
                        twWriter.Write(row[col].ToString().Replace("<b>", string.Empty).Replace("</b>", string.Empty));
                        twWriter.Write(seperator);
                    }
                    twWriter.WriteLine();
                }
                twWriter.WriteLine();
                // this code is for totals....................
                foreach (string col in column)
                {

                    double footer = 0;

                    if (Totals.Where(t => t.ColumnName.Contains(col)).Any())
                    {

                        switch (Totals.Where(t => t.ColumnName.Contains(col)).Select(g => g.Totaltype).SingleOrDefault())
                        {
                            case "Count":
                                {
                                    footer = dt.AsEnumerable().Count();
                                    twWriter.Write("Count" + ":" + footer.ToString());
                                    break;
                                }
                            case "Sum":
                                {
                                    if (col.ToLower().Contains("cost"))
                                    { twWriter.Write("Total" + ":" + dt.AsEnumerable().Sum(s => s.Field<decimal>(col)).ToString()); }
                                    else
                                    {
                                        footer = dt.AsEnumerable().Sum(s => s.Field<int>(col));
                                        twWriter.Write("Total" + ":" + footer.ToString());
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
                                    twWriter.Write("Total" + ":" + t1.ToString());
                                    break;
                                }
                            case "Avg":
                                {
                                    footer = dt.AsEnumerable().Average(s => s.Field<int>(col));
                                    twWriter.Write("Avg" + ":" + footer.ToString());
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
                                    twWriter.Write("Avg" + ":" + t1.ToString());
                                    break;
                                }
                        }
                    }
                    twWriter.Write(seperator);
                }



            }
            catch { 
            
            
            
            }
        
        
        
        }
        public static void Departmentalreport(TextWriter twWriter, string seperator, DataSet dt, List<string> columns, string level, int direclevel, List<ColumnTotals> Totals)
        {
            try
            {

                if (level == "4")
                {
                    GetStringTable(twWriter, dt.Tables[2], 0, seperator, Totals);
                    twWriter.WriteLine();
                    if (direclevel != 1)
                    {
                        GetStringTable(twWriter, dt.Tables[direclevel + 2], 1, seperator, Totals);
                    }
                    else
                    {
                        GetStringTable(twWriter, dt.Tables[4], 1, seperator, Totals);
                        GetExtlevelTable(twWriter, dt.Tables[0], dt.Tables[3], seperator, Totals);
                    }
                }
                else
                {
                    if (direclevel == 1)
                    {
                        GetStringTable(twWriter, dt.Tables[3], 0, seperator, Totals);
                        twWriter.WriteLine();
                    }
                    else if (direclevel == 2)
                    {
                        GetStringTable(twWriter, dt.Tables[2], 0, seperator, Totals);
                        twWriter.WriteLine();
                        GetStringTable(twWriter, dt.Tables[4], 0, seperator, Totals);
                        twWriter.WriteLine();
                        GetExtlevelTable(twWriter, dt.Tables[0], dt.Tables[3], seperator, Totals);
                    }
                    else if (direclevel == 3)
                    {
                        GetStringTable(twWriter, dt.Tables[2], 0, seperator, Totals);
                        twWriter.WriteLine();
                        GetStringTable(twWriter, dt.Tables[5], 1, seperator, Totals);
                        twWriter.WriteLine();
                        GetStringTable(twWriter, dt.Tables[4], 1, seperator, Totals);
                        twWriter.WriteLine();
                        GetExtlevelTable(twWriter, dt.Tables[0], dt.Tables[3], seperator, Totals);
                    }
                    else if (direclevel == 4)
                    {
                        GetStringTable(twWriter, dt.Tables[2], 0, seperator, Totals);
                        twWriter.WriteLine();
                        GetStringTable(twWriter, dt.Tables[6], 1, seperator, Totals);
                        twWriter.WriteLine();
                        GetStringTable(twWriter, dt.Tables[5], 1, seperator, Totals);
                        twWriter.WriteLine();
                        GetStringTable(twWriter, dt.Tables[4], 1, seperator, Totals);
                        twWriter.WriteLine();
                        GetExtlevelTable(twWriter, dt.Tables[0], dt.Tables[3], seperator, Totals);
                    }
                }

            }
            catch { }
        }
        public static void GetStringTable(TextWriter twWriter, DataTable dt, int level, string seperator, List<ColumnTotals> Totals)
        {


            List<string> column = new List<string>();
            //creating table headers
            foreach (DataColumn colu in dt.Columns)
            {
                if (colu.ColumnName != "Department" && colu.ColumnName != "Totalnonform" && colu.ColumnName != "indurnonformated" && colu.ColumnName != "outdurnonformated")
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
            foreach (string col in column)
            {

                double footer = 0;

                if (Totals.Where(t => t.ColumnName.Contains(col)).Any())
                {

                    switch (Totals.Where(t => t.ColumnName.Contains(col)).Select(g => g.Totaltype).SingleOrDefault())
                    {
                        case "Count":
                            {
                                footer = dt.AsEnumerable().Count();
                                twWriter.Write("Count" + ":" + footer.ToString());
                                break;
                            }
                        case "Sum":
                            {
                                if (col.ToLower().Contains("cost"))
                                { twWriter.Write("Total" + ":" + dt.AsEnumerable().Sum(s => s.Field<decimal>(col)).ToString()); }
                                else
                                {
                                    footer = dt.AsEnumerable().Sum(s => s.Field<int>(col));
                                    twWriter.Write("Total" + ":" + footer.ToString());
                                }


                                break;
                            }
                        case "FormatedSum":
                            {
                                string calc = "";
                                switch (col)
                                {
                                    case "Incoming Duration":
                                        calc = "indurnonformated";
                                        break;
                                    case "Outgoing Duration":
                                        calc = "outdurnonformated";
                                        break;
                                    case "Total Duration":
                                        calc = "Totalnonform";
                                        break;

                                }
                                int dis = dt.AsEnumerable().Sum(s => s.Field<int>(calc));

                                TimeSpan t1 = new TimeSpan(0, 0, dis);
                                twWriter.Write("Total" + ":" + t1.ToString());
                                break;
                            }
                        case "Avg":
                            {
                                footer = dt.AsEnumerable().Average(s => s.Field<int>(col));
                                twWriter.Write("Avg" + ":" + footer.ToString());
                                break;
                            }

                    }
                }
                twWriter.Write(seperator);
            }
        }
        public static void GetExtlevelTable(TextWriter twWriter, DataTable des, DataTable dt, string seperator, List<ColumnTotals> Totals)
        {

            try
            {

                var depts = dt.AsEnumerable().Select(s => s.Field<string>("Cost Centre")).Distinct();

                foreach (var v in depts)
                {
                    twWriter.WriteLine();
                    twWriter.Write(v.ToString());
                    twWriter.WriteLine(" ");

                    twWriter.WriteLine();

                    foreach (DataColumn column in des.Columns)
                    {
                        if (column.ColumnName != "Department" && column.ColumnName != "Totalnonform" && column.ColumnName != "indurnonformated" && column.ColumnName != "outdurnonformated")
                        {
                            twWriter.Write(column.ColumnName);
                            twWriter.Write(seperator);
                        }

                    }


                    var dest = from db in des.AsEnumerable()
                               where db.Field<string>("Department") == v.ToString()
                               select new
                               {
                                   Destinationname = db.Field<object>("Destination Name"),
                                   Totalcalls = db.Field<object>("Total Calls"),
                                   Totalduration = db.Field<object>("Total Duration"),
                                   Totalnonform = db.Field<object>("Totalnonform"),
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
                    #region "Problem with totals"
                    try
                    {
                        foreach (DataColumn colad in des.Columns)
                        {
                            string col = colad.ColumnName;
                            double footer = 0;
                            if (col != "Department" && col != "Totalnonform")
                            {
                                if (Totals.Where(t => t.ColumnName.Contains(col)).Any())
                                {

                                    switch (Totals.Where(t => t.ColumnName.Contains(col)).Select(g => g.Totaltype).SingleOrDefault())
                                    {
                                        case "Count":
                                            {
                                                footer = dest.AsEnumerable().Count();
                                                twWriter.Write("Count" + ":" + footer.ToString());
                                                break;
                                            }
                                        case "Sum":
                                            {
                                                if (col.ToLower().Contains("cost"))
                                                { twWriter.Write("Total" + ":" + dest.AsEnumerable().Sum(s => Convert.ToDouble(s.Cost))); }
                                                else
                                                {
                                                    footer = dest.AsEnumerable().Sum(s => Convert.ToInt32(s.Totalcalls));
                                                    twWriter.Write("Total" + ":" + footer.ToString());
                                                }


                                                break;
                                            }
                                        case "FormatedSum":
                                            {

                                                int form = dest.AsEnumerable().Sum(s => Convert.ToInt32(s.Totalnonform));
                                                TimeSpan t1 = new TimeSpan(0, 0, form);
                                                twWriter.Write("Total" + ":" + t1.ToString());
                                                break;
                                            }


                                    }
                                }
                                twWriter.Write(seperator);
                            }
                        }
                    }
                    catch { }
                    #endregion
                    twWriter.WriteLine();
                    foreach (DataColumn column in dt.Columns)
                    {
                        if (column.ColumnName != "Cost Centre" && column.ColumnName != "indurnonformated" && column.ColumnName != "outdurnonformated")
                        {
                            twWriter.Write(column.ColumnName);
                            twWriter.Write(seperator);
                        }


                    }


                    twWriter.WriteLine(" ");
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
                        twWriter.WriteLine();
                        twWriter.Write(all.Name.ToString().Replace(" ", string.Empty));
                        twWriter.Write(seperator);
                        twWriter.Write(all.Extension.ToString().Replace(" ", string.Empty));
                        twWriter.Write(seperator);
                        twWriter.Write(all.OutgoingCalls.ToString().Replace(" ", string.Empty));
                        twWriter.Write(seperator);
                        twWriter.Write(all.OutgoingDuration.ToString().Replace(" ", string.Empty));
                        twWriter.Write(seperator);
                        twWriter.Write(all.RingResponse.ToString().Replace(" ", string.Empty));
                        twWriter.Write(seperator);
                        twWriter.Write(all.AbandonedCalls.ToString().Replace(" ", string.Empty));
                        twWriter.Write(seperator);
                        twWriter.Write(all.IncomingCalls.ToString().Replace(" ", string.Empty));
                        twWriter.Write(seperator);
                        twWriter.Write(all.IncomingDuration.ToString().Replace(" ", string.Empty));
                        twWriter.Write(seperator);
                        twWriter.Write(all.Cost.ToString().Replace(" ", string.Empty));

                    }
                    # region "Problem with totals...."
                    twWriter.WriteLine(" ");
                    try
                    {
                        foreach (DataColumn colnm in dt.Columns)
                        {
                            string col = colnm.ColumnName;
                            string footer = "";
                            if (col != "Cost Centre" && col != "indurnonformated" && col != "outdurnonformated")
                            {
                                if (Totals.Where(t => t.ColumnName.Contains(col)).Any())
                                {

                                    switch (Totals.Where(t => t.ColumnName.Contains(col)).Select(g => g.Totaltype).SingleOrDefault())
                                    {
                                        case "Count":
                                            {
                                                footer = records.AsEnumerable().Count().ToString();
                                                twWriter.Write("Count" + ":" + footer.ToString());
                                                break;
                                            }
                                        case "Sum":
                                            {
                                                if (col.ToLower().Contains("cost"))
                                                { twWriter.Write("Total" + ":" + records.AsEnumerable().Sum(s => Convert.ToDouble(s.Cost)).ToString()); }


                                                else if (col.Contains("Incoming Calls"))
                                                {
                                                    footer = Convert.ToString(records.AsEnumerable().Sum(s => Convert.ToInt32(s.IncomingCalls)));
                                                    twWriter.Write("Total" + ":" + footer.ToString());
                                                }
                                                else if (col.Contains("Outgoing Calls"))
                                                {
                                                    footer = Convert.ToString(records.AsEnumerable().Sum(s => Convert.ToInt32(s.OutgoingCalls)));
                                                    twWriter.Write("Total" + ":" + footer.ToString());
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
                                                twWriter.Write("Total" + ":" + t1.ToString());
                                                break;
                                               
                                            }
                                        case "Avg":
                                            {
                                                footer = Convert.ToString(Math.Round( records.AsEnumerable().Average(s => Convert.ToInt32(s.RingResponse)),2));
                                                twWriter.Write("Avg" + ":" + footer.ToString());
                                                break;
                                            }

                                    }
                                }
                                twWriter.Write(seperator);
                            }

                        }
                    }
                    catch { }
                    #endregion
                    twWriter.WriteLine(" ");


                }
            }
            catch { }


        }
    }
}
