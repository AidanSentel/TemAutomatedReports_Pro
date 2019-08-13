using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace TEMAutomatedReports
{
    class KPI
    {
        public string KpiName { get; set; }
        public string KPiOperation { get; set; }
        public string Columnnames { get; set; }
        public static string GenerateKPIReport(DataSet ds, string report, int UserId)
        {
           IEnumerable<tbl_KPI> SlA = Datamethods.GetKPI_Site(UserId);
            List<KPI> kpi =  Datamethods.GetKPI_ForReport(report, "Kpi");
            StringBuilder HTMLtoRender = new StringBuilder();
            try
            {
                HTMLtoRender.Append("<p><b>KPI Vs SLA:</b></P>");
                HTMLtoRender.Append("<table width='100%' style='border-style: dotted'>");
                foreach (KPI k in kpi)
                {
                    HTMLtoRender.Append("<tr>");
                    HTMLtoRender.Append("<td colspan='3'>");
                    HTMLtoRender.Append("<CENTER><h3>" + k.KpiName + "</h3></CENTER>");
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.Append("</tr>");
                    HTMLtoRender.Append("<tr>");
                    HTMLtoRender.Append("<td style='width:200px'>");
                    HTMLtoRender.Append("SLA =<b>" + GetSlarate(k.KpiName, SlA) + "</b>");
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.Append("<td style='width:200px'>");
                    double main = ds.Tables[0].AsEnumerable().Sum(s => s.Field<int>(k.Columnnames.Split(',')[0]));
                    double kp = ds.Tables[0].AsEnumerable().Sum(s => s.Field<int>(k.Columnnames.Split(',')[1]));
                    double ans = Math.Round(((kp - main) / kp) * 100, 2);
                    HTMLtoRender.Append("KPI =<b>" + Math.Round( (100 - ans),2).ToString() + "</b>%");
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.Append("<td style='width:100px'>");
                    HTMLtoRender.Append("Diff =<b>" + ans.ToString() + "</b>%");
                    HTMLtoRender.Append("</td>");
                    HTMLtoRender.Append("</tr>");
                }
                HTMLtoRender.Append("</table>");
            }
            catch { }
            return HTMLtoRender.ToString();
        
        }

        private static string GetSlarate(string Kpiname, IEnumerable<tbl_KPI> SlA)
        {
            string value = "";
            foreach (tbl_KPI k1 in SlA)
            {
                switch (Kpiname)
                { 
                    case "Response Rate":
                        value = "100% in " + k1.KPI_Response_Rate_Day.ToString()+" Sec";
                        break;
                    case "Answered Rate":
                        value = k1.KPI_Answered_Rate_Day.ToString() + " % Calls";
                        break;
                        
                    case "Abandoned Rate":
                        value =  k1.KPI_Abandoned_Rate_Day.ToString() + " % Calls";
                        break;
                
                }
                
            }
            return value;
           
        }
       


    }
}
