using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Drawing;

namespace TEMAutomatedReports
{
    class Header
    {
       public static Color[] colurs = new Color[]
            {
             System.Drawing.Color.MidnightBlue,             //dashboard
             System.Drawing.Color.LightSeaGreen ,           //Landline
             System.Drawing.Color.DarkRed,                //Tariff
             System.Drawing.Color.Goldenrod,                //assert
             System.Drawing.Color.DarkSalmon,               //bill val
             System.Drawing.Color.OrangeRed,                //KPI
             System.Drawing.Color.Crimson,                  //Inventory
             System.Drawing.Color.Gold,                     //mobile
             System.Drawing.Color.DarkRed ,
             System.Drawing.Color.Cyan ,                    //Admin
             System.Drawing.Color.Tomato,                   //Trend
             System.Drawing.Color.MediumSeaGreen          //Reporting
            };
        public static string  Createheader(DataSet ds , string report , string section)
        {
            List<ColumnTotals> totals = Datamethods.GetHeaderValues(report, section);
            int tables = ds.Tables.Count;
            StringBuilder HTMLtoRender = new StringBuilder("<table class='table table-striped' align ='center' width='100%' style=\"font:15px/20px arial,sans-serif\">", 100000);
            try
            {
                HTMLtoRender.Append("<tr>");
                int i = 0;
                foreach (ColumnTotals c1 in totals)
                {
                    foreach (DataTable dt in ds.Tables)
                    {
                        if (dt.Columns.Contains(c1.ColumnName.Split('~')[0]))
                        {
                            if (!HTMLtoRender.ToString().Contains(c1.Friendlyname))
                            {
                                Color color = colurs[i];
                                string myHexString = String.Format("#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B);
                                HTMLtoRender.Append("<td style='width:100px;height:50px;color:whitesmoke;background-color:" + myHexString + " '>");
                                HTMLtoRender.Append("<CENTER><h2>" + c1.Friendlyname + "</h2></CENTER>");
                                switch (c1.Totaltype)
                                {
                                    case "Sum":
                                        if (c1.ColumnName.ToLower().Contains("cost"))
                                        { HTMLtoRender.Append("</br><CENTER><h2>" + Math.Round(dt.AsEnumerable().Sum(s => s.Field<decimal>(c1.ColumnName)), 2).ToString() + "</h2></CENTER>"); }
                                        else
                                        {
                                            HTMLtoRender.Append("</br><CENTER><h2>" + dt.AsEnumerable().Sum(s => s.Field<int>(c1.ColumnName)).ToString() + "</h2></CENTER>");
                                        }
                                        break;
                                    case "Field":
                                        HTMLtoRender.Append("</br><CENTER><h2>" + Convert.ToString(dt.Rows[0][c1.ColumnName]) + "</h2></CENTER>");
                                        break;
                                    case "Minus":
                                        int toatalminus = 0;
                                        if (dt.Columns.Contains(c1.ColumnName.Split('~')[0]))
                                        {
                                            toatalminus = dt.AsEnumerable().Sum(s => s.Field<int>(c1.ColumnName.Split('~')[0]));
                                        }
                                        if (dt.Columns.Contains(c1.ColumnName.Split('~')[1]))
                                        {
                                            toatalminus = toatalminus - dt.AsEnumerable().Sum(s => s.Field<int>(c1.ColumnName.Split('~')[1]));
                                        }
                                        HTMLtoRender.Append("</br><CENTER><h2>" + Convert.ToString(toatalminus) + "</h2></CENTER>");
                                        break;
                                    case "Percentage1":
                                        int percent1 = 0, percent2 = 1;
                                        if (dt.Columns.Contains(c1.ColumnName.Split('~')[0]))
                                        {
                                            percent1 = dt.AsEnumerable().Sum(s => s.Field<int>(c1.ColumnName.Split('~')[0]));
                                        }
                                        if (dt.Columns.Contains(c1.ColumnName.Split('~')[1]))
                                        {
                                            percent2 = dt.AsEnumerable().Sum(s => s.Field<int>(c1.ColumnName.Split('~')[1]));
                                        }

                                        if (percent1 != 0 && percent2 != 0)
                                        {
                                            HTMLtoRender.Append("</br><CENTER><h2>" + Math.Round((Convert.ToDouble(percent1) / Convert.ToDouble(percent2)) * 100, 2).ToString() + "</h2></CENTER>");
                                        }
                                        break;
                                    case "Percentage2":
                                        int percent11 = 0, percent22 = 1;
                                        if (dt.Columns.Contains(c1.ColumnName.Split('~')[0]))
                                        {
                                            percent11 = dt.AsEnumerable().Sum(s => s.Field<int>(c1.ColumnName.Split('~')[0]));
                                        }
                                        if (dt.Columns.Contains(c1.ColumnName.Split('~')[1]))
                                        {
                                            percent22 = dt.AsEnumerable().Sum(s => s.Field<int>(c1.ColumnName.Split('~')[1]));
                                        }

                                        if (percent11 != 0 && percent22 != 0)
                                        {
                                            HTMLtoRender.Append("</br><CENTER><h2>" + (100 - Math.Round((Convert.ToDouble(percent11) / Convert.ToDouble(percent22)) * 100, 2)).ToString() + "</h2></CENTER>");
                                        }
                                        break;
                                    case "Count":
                                        HTMLtoRender.Append("</br><CENTER><h2>" + dt.Rows.Count.ToString() + "</h2></CENTER>");
                                        break;

                                    case "Avg":
                                        HTMLtoRender.Append("</br><CENTER><h2>" + Math.Round(dt.AsEnumerable().Average(s => s.Field<int>(c1.ColumnName)), 2).ToString() + "</h2></CENTER>");

                                        break;

                                    case "Avg Duration":
                                        int val = Convert.ToInt32(Math.Round(dt.AsEnumerable().Average(s => s.Field<int>(c1.ColumnName))));
                                        TimeSpan ts = new TimeSpan(0, 0, val);
                                        HTMLtoRender.Append("</br><CENTER><h2>" + ts.ToString() + "</h2></CENTER>");
                                        break;
                                    case "Total Duration":
                                        if (c1.ColumnName.Contains("~"))
                                        {
                                            int toatalsum = 0;
                                            foreach (string str in c1.ColumnName.Split('~'))
                                            {
                                                toatalsum += Convert.ToInt32(dt.AsEnumerable().Sum(s => s.Field<int>(str)));
                                            }
                                            TimeSpan ts1 = new TimeSpan(0, 0, toatalsum);
                                            HTMLtoRender.Append("</br><CENTER><h2>" + ts1.ToString() + "</h2></CENTER>");
                                        }
                                        else
                                        {
                                            int val1 = Convert.ToInt32(dt.AsEnumerable().Sum(s => s.Field<int>(c1.ColumnName)));
                                            TimeSpan ts1 = new TimeSpan(0, 0, val1);
                                            HTMLtoRender.Append("</br><CENTER><h2>" + ts1.ToString() + "</h2></CENTER>");
                                        }
                                        break;
                                    default:
                                        break;

                                }
                                HTMLtoRender.Append("</td>");
                                i++;
                            }
                        }
                        else
                        {
                            if (c1.Friendlyname == "BREAK")
                            {
                                HTMLtoRender.Append("</tr>");
                                HTMLtoRender.Append("<tr>");
                                break;
                            }
                        }

                    }


                }
                HTMLtoRender.Append("</tr>");
            }
            catch { }
            HTMLtoRender.Append("</table>");
            return HTMLtoRender.ToString();
        
        }

        public static string Secondheader(DataSet ds, string report, string section)
        {
            List<KPI> totals = Datamethods.Getsecondtotals_ForReport(report, section);
            int tables = ds.Tables.Count;
            StringBuilder HTMLtoRender = new StringBuilder() ;
            if (totals.Count > 0)
            {
                HTMLtoRender.AppendFormat("<table class='table table-striped' align ='center' width='100%' style=\"font:15px/20px arial,sans-serif\">", 100000);
                HTMLtoRender.Append("<tr>");
                foreach (KPI kp in totals)
                {
                    foreach (DataTable dt in ds.Tables)
                    {
                        if (dt.Columns.Contains(kp.Columnnames.Split(',')[0]))
                        {
                            double main = ds.Tables[0].AsEnumerable().Sum(s => s.Field<int>(kp.Columnnames.Split(',')[0]));
                            double pp = ds.Tables[0].AsEnumerable().Sum(s => s.Field<int>(kp.Columnnames.Split(',')[1]));
                            double ans = Math.Round(((pp - main) / pp) * 100, 2);
                            HTMLtoRender.Append("<td style='width:100px;height:50px;color:whitesmoke;background-color:gray'>");
                            HTMLtoRender.Append("<CENTER><b>" + kp.KpiName + "</b></CENTER>");
                            HTMLtoRender.Append("</br><CENTER><b>"+(100 - ans).ToString() + "</b></CENTER>");
                            HTMLtoRender.Append("</td>");
                        }                  
                    
                    }
                
                }

                HTMLtoRender.Append("</tr>");
                HTMLtoRender.Append("</table>");
            }
           
           
            return HTMLtoRender.ToString();
        
        }

    }
}
