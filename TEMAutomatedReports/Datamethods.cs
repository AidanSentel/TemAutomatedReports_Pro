using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DateTimeExtensions;
using System.Globalization;

namespace TEMAutomatedReports
{
    class Datamethods : IDisposable
    {
        /* 
         This class has all database operation methods 
          
         */
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

        public static int Starttime { get; set; }
        // This method will get and return all the schedules based on report frequency...
        internal static IQueryable<Schedule> GetReportsData(string frequency)
        {

            // for testing

            List<int> ids = new List<int>();


            //      ids.Add(39);
            //      ids.Add(43);
            //ids.Add(44);
            //ids.Add(45);

            //ids.Add(2321);
            //            
            IQueryable<Schedule> reports = from rep in DbContext.tbl_automatedreports.AsQueryable()
                                           join sp in DbContext.tbl_reports on rep.schedule_storedprocedure equals sp.reports_storedprocedure
                                           where rep.schedule_frequency == frequency
                                           && rep.schedule_Active == true

                                           // && rep.schedule_frequency == "quarterly" 
                                           // && rep.schedule_reportformat != "html"       

                                           // && rep.schedule_filters.Contains("2014-02-24")
                                           ///&& rep.schedule_id_PK == 1514
                                           //&& rep.schedule_id_PK == 1391

                                           //  && ids.Contains( rep.schedule_id_PK)//. (98,426,596)

                                           //     && rep.schedule_filters.Contains("2019-06-01 ")
                                           select new Schedule
                                           {
                                               ID = rep.schedule_id_PK,
                                               ListofFilters = rep.schedule_filters,
                                               EmailAddresses = rep.schedule_emailaddresses,
                                               StoredProcedureName = rep.schedule_storedprocedure,
                                               Columns = sp.reports_columns,
                                               CreatedDate = rep.schedule_createddate,
                                               Frequency = rep.schedule_frequency,
                                               Chosennodelist = rep.schedule_chosenNodeIds,
                                               ReportName = rep.schedule_reportname,
                                               Selectedname = rep.schedule_schedulename,
                                               Type = rep.schedule_reportformat,
                                               Totals = sp.reports_columntotal,
                                               Portfolioreportid = rep.Protfolio_ID ?? 0,
                                               Status = rep.schedule_Active ?? true,
                                               Time = rep.schedule_Time ?? 0,
                                               GraphBindings = sp.reports_graphbindings ?? string.Empty,
                                               GraphType = rep.schedule_Graph ?? string.Empty,
                                               ReportingSection = rep.schedule_reporttype,
                                               UserId = rep.schedule_user_id_FK ?? 0,
                                               GraphHeaders = sp.reports_totals ?? string.Empty
                                           };

            switch (frequency)
            {   // To get weekly reports based on day of week
                case "fixed":
                    reports = reports.Where(s => s.Time == DateTime.Now.TimeOfDay.Hours);
                    break;
                case "weekly":
                    reports = reports.Where(s => s.CreatedDate.DayOfWeek == DateTime.Now.DayOfWeek - 1);
                    break;
                // to get monthly reports based on current date
                case "monthly":
                    reports = reports.Where(s => s.CreatedDate.Day == DateTime.Now.Day);
                    //reports = reports.Where(s => s.CreatedDate == new DateTime(2016,03,02,00,00,00)); // DateTime.Now.Day);
                    break;
                // to get monthly report based on company financial year..
                case "FixedMonth":
                    reports = reports.Where(s => s.CreatedDate.Day == DateTime.Now.Day);
                    break;
                // to get quaterly reports based on date
                case "quarterly":
                    // need to update all quartely reports createated data to three month
                    reports = reports.Where(s => s.CreatedDate.Date == DateTime.Now.Date);
                    break;
                case "yearly":
                    reports = reports.Where(s => s.CreatedDate.Day == DateTime.Now.Day && s.CreatedDate.Month == DateTime.Now.Month);
                    break;

            }

            GC.Collect();
            return reports;

        }
        //This method will get all report frequncies form DB
        internal static IQueryable<tbl_reportfrequency> Frequency()
        {
            int sTime = Datamethods.Starttime;
            sTime = 2;
            IQueryable<tbl_reportfrequency> frequen = from freq in DbContext.tbl_reportfrequencies
                                                      where freq.Frequency_Time == sTime   //DateTime.Now.TimeOfDay.Hours     //where freq.Frequency_Time == 3//DateTime.Now.TimeOfDay.Hours
                                                      select freq;
            return frequen;


        }
        public static void Updatetimes(tbl_reportfrequency freq)
        {

            if (Starttime > 8 && Starttime < 18)
            {
                freq.Frequency_Time = Starttime + 1;
            }
            else if (Starttime == 18) { freq.Frequency_Time = 9; }
            DbContext.SubmitChanges();
        }

        // this method will update call records based on report frequncy........
        public static void Updatereport(Schedule report)
        {
            try
            {
                DateTime From = DateTime.Now, To = DateTime.Now;

                try
                {
                    From = Convert.ToDateTime(report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\''));
                    To = Convert.ToDateTime(report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\''));
                }
                catch { }
                string inter = "";
                switch (report.Frequency)
                {
                    case "hourly":

                        if (Starttime != 17)
                        {
                            inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddHours(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                            report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddHours(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        }
                        else
                        {
                            inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddHours(16).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                            report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddHours(16).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        }
                        break;

                    case "fixed":
                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        break;
                    case "FixedMonth":
                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + CovidienMonthstartdate(From).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + CovidienMonthenddate(To).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.CreatedDate = CovidienMonthenddate(To).AddDays(1);
                        break;

                    case "daily":
                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        break;
                    case "MonthToDt":
                        if (report.CreatedDate.Day == To.AddDays(1).Day)
                        {
                            inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddMonths(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                            report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        }
                        else
                        {
                            report.ListofFilters = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        }

                        break;
                    case "weekly":
                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddDays(7).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddDays(7).ToString("yyyy-MM-dd HH:mm:ss") + "'");

                        break;
                    case "monthly":
                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddMonths(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + From.AddMonths(2).AddDays(-1).ToString("yyyy-MM-dd 23:59:59") + "'");

                        break;
                    case "quarterly":
                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddMonths(3).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + From.AddMonths(3).AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.CreatedDate = report.CreatedDate.AddMonths(3);
                        break;
                    case "yearly":
                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddYears(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddYears(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");


                        break;

                }

                // Need an update statement here.....

                //           var rep = DbContext.tbl_automatedreports.Where(s => s.schedule_id_PK == report.ID).SingleOrDefault();
                //          rep.schedule_filters = report.ListofFilters;
                //         rep.schedule_createddate = report.CreatedDate;
                //        DbContext.SubmitChanges();

            }
            catch { GenerateReports.ReportStatus(report.ID, "Erro Updating Report"); }
        }

        public static void Undoreport(Schedule report)
        {
            try
            {
                DateTime From = DateTime.Now, To = DateTime.Now;

                try
                {
                    From = Convert.ToDateTime(report.ListofFilters.Split('=', ',')[5].TrimStart('\'').TrimEnd('\''));
                    To = Convert.ToDateTime(report.ListofFilters.Split('=', ',')[7].TrimStart('\'').TrimEnd('\''));

                }
                catch { }
                string inter = "";
                switch (report.Frequency)
                {

                    case "hourly":

                        if (Starttime != 17)
                        {
                            inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddHours(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                            report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddHours(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        }
                        else
                        {
                            inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddHours(16).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                            report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddHours(16).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        }
                        break;

                    case "fixed":

                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddDays(1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        break;

                    case "daily":

                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        break;
                    case "weekly":
                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddDays(-7).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddDays(-7).ToString("yyyy-MM-dd HH:mm:ss") + "'");

                        break;
                    case "monthly":
                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddMonths(-1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + From.AddMonths(-2).AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss") + "'");

                        break;
                    case "quaterly":
                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddMonths(-3).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + From.AddMonths(-3).AddDays(-1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.CreatedDate = report.CreatedDate.AddMonths(3);
                        break;
                    case "yearly":
                        inter = report.ListofFilters.Replace(report.ListofFilters.Split('=', ',')[5], "'" + From.AddYears(-1).ToString("yyyy-MM-dd HH:mm:ss") + "'");
                        report.ListofFilters = inter.Replace(report.ListofFilters.Split('=', ',')[7], "'" + To.AddYears(-1).ToString("yyyy-MM-dd HH:mm:ss") + "'");


                        break;

                }

                // Need an update statement here.....

                var rep = DbContext.tbl_automatedreports.Where(s => s.schedule_id_PK == report.ID).SingleOrDefault();
                rep.schedule_filters = report.ListofFilters;
                rep.schedule_createddate = report.CreatedDate;
                DbContext.SubmitChanges();


            }
            catch { GenerateReports.ReportStatus(report.ID, "Erro Updating Report"); }
        }
        public static DateTime CovidienMonthstartdate(DateTime date)
        {
            int[] list = new int[] { 0, 4, 9, 13, 17, 22, 26, 30, 35, 39, 43, 48, 52 };
            DateTime firstDayOfyear = FirstDateOfWeek((date.Year), 40, CultureInfo.CurrentCulture).AddDays(-1);
            if (GetIso8601WeekOfYear(date) >= 1 && (GetIso8601WeekOfYear(date) < 39))
            {
                firstDayOfyear = FirstDateOfWeek((date.Year - 1), 40, CultureInfo.CurrentCulture).AddDays(-1);
            }
            DateTime tobereturn = firstDayOfyear;
            for (int i = 0; i < 12; i++)
            {
                // tobereturn = Covidiendateupdate(firstDayOfyear, i);
                if (tobereturn.AddDays(list[i] * 7) == date)
                {
                    tobereturn = firstDayOfyear.AddDays(list[i + 1] * 7);
                    break;
                }


            }
            return tobereturn;
        }
        // will return covidien mont ending date
        public static DateTime CovidienMonthenddate(DateTime date)
        {
            int[] list = new int[] { 0, 4, 9, 13, 17, 22, 26, 30, 35, 39, 43, 48, 52 };
            DateTime firstDayOfyear = FirstDateOfWeek((date.Year), 40, CultureInfo.CurrentCulture).AddDays(-1);
            if (GetIso8601WeekOfYear(date) >= 1 && (GetIso8601WeekOfYear(date) < 38))
            {
                firstDayOfyear = FirstDateOfWeek((date.Year - 1), 40, CultureInfo.CurrentCulture).AddDays(-1);
            }
            DateTime lasttDayOfyear = firstDayOfyear.AddDays(-1).SetTime(23, 59, 59);
            DateTime tobereturn = firstDayOfyear;
            for (int i = 0; i < 12; i++)
            {
                //tobereturn = Covidiendateupdate(firstDayOfyear, i);
                if (lasttDayOfyear.AddDays(list[i] * 7) == date)
                {
                    tobereturn = lasttDayOfyear.AddDays(list[i + 1] * 7);
                    break;
                }

            }
            return tobereturn;
        }
        // to get date based on week...
        public static DateTime FirstDateOfWeek(int year, int weekOfYear, System.Globalization.CultureInfo ci)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = (int)ci.DateTimeFormat.FirstDayOfWeek - (int)jan1.DayOfWeek;
            DateTime firstWeekDay = jan1.AddDays(daysOffset);
            int firstWeek = ci.Calendar.GetWeekOfYear(jan1, ci.DateTimeFormat.CalendarWeekRule, ci.DateTimeFormat.FirstDayOfWeek);
            if (firstWeek <= 1 || firstWeek > 50)
            {
                weekOfYear -= 1;
            }
            return firstWeekDay.AddDays(weekOfYear * 7);
        }
        // TO get week number based on date
        public static int GetIso8601WeekOfYear(DateTime time)
        {
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }

            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        private static void Update(string Filters, string frequency)
        {


            switch (frequency)
            {
                case "daily":
                    break;
                case "weekly":
                    break;
                case "monthly":
                    break;
                case "quaterly":
                    break;
                case "yearly":
                    break;

            }


        }
        // this method will get all the header totals required for KPI and dashboard
        public static List<ColumnTotals> GetHeaderValues(string name, string section)
        {
            IEnumerable<tbl_ReportTotal> reptotals = DbContext.tbl_ReportTotals.Where(s => s.Report_Name == name & s.Report_Type == section);

            List<ColumnTotals> totals = new List<ColumnTotals>();
            try
            {
                foreach (tbl_ReportTotal t1 in reptotals)
                {
                    foreach (string individual in t1.Report_Totals.ToString().Split(','))
                    {
                        ColumnTotals c1 = new ColumnTotals();
                        string[] seperator = individual.Split('@', '#');
                        c1.Friendlyname = seperator[0];
                        c1.Totaltype = seperator[1];
                        c1.ColumnName = seperator[2];
                        totals.Add(c1);

                    }

                }
            }
            catch { }
            return totals;
        }
        // this report will get all KPi params
        public static IEnumerable<tbl_KPI> GetKPI_Site(int UserId)
        {

            List<int> site = DbContext.tbl_usersites.Where(s => s.usersite_user_id_FK == UserId).Select(s => s.usersite_site_id_FK).ToList();
            IEnumerable<tbl_KPI> kpi = DbContext.tbl_KPIs.Where(s => site.Contains(s.site_id_PK ?? 0));
            return kpi;


        }
        // this method will get list of kpis indicators for a report
        public static List<KPI> GetKPI_ForReport(string name, string section)
        {
            IEnumerable<tbl_ReportTotal> reptotals = DbContext.tbl_ReportTotals.Where(s => s.Report_Name == name & s.Report_Type == section);

            List<KPI> KPi = new List<KPI>();
            try
            {
                foreach (tbl_ReportTotal t1 in reptotals)
                {
                    foreach (string individual in t1.Report_Performance.ToString().Split('-'))
                    {
                        KPI c1 = new KPI();
                        string[] seperator = individual.Split('@', '#');
                        c1.KpiName = seperator[0];
                        c1.KPiOperation = seperator[1];
                        c1.Columnnames = seperator[2];
                        KPi.Add(c1);

                    }

                }
            }
            catch { }
            return KPi;


        }
        // this method will get graph headers and column names
        public static List<GraphHeaders> GetGraphHeadernames(string names)
        {
            List<GraphHeaders> headers = new System.Collections.Generic.List<GraphHeaders>();
            try
            {
                foreach (string s in names.Split(','))
                {
                    GraphHeaders g1 = new GraphHeaders();
                    g1.Headername = s.Split('@')[0];
                    g1.ColumnName = s.Split('@')[1];
                    headers.Add(g1);
                }
            }
            catch { }
            return headers;
        }
        // this method will get list of second headers for a report
        public static List<KPI> Getsecondtotals_ForReport(string name, string section)
        {
            IEnumerable<tbl_ReportTotal> reptotals = DbContext.tbl_ReportTotals.Where(s => s.Report_Name == name & s.Report_Type == section);

            List<KPI> KPi = new List<KPI>();
            foreach (tbl_ReportTotal t1 in reptotals)
            {
                foreach (string individual in t1.Report_SubTotals.ToString().Split('-'))
                {
                    KPI c1 = new KPI();
                    string[] seperator = individual.Split('@', '#');
                    c1.KpiName = seperator[0];
                    c1.KPiOperation = seperator[1];
                    c1.Columnnames = seperator[2];
                    KPi.Add(c1);

                }

            }
            return KPi;


        }

        // in case of tables need to be joined with common columns.. this method will retun info about that
        public static List<TableJoin> Table_Joins(string name, string section, out bool hasvalue)
        {
            IEnumerable<tbl_ReportTotal> reptotals = DbContext.tbl_ReportTotals.Where(s => s.Report_Name == name & s.Report_Type == section);
            List<TableJoin> operations = new List<TableJoin>();
            hasvalue = false;
            foreach (tbl_ReportTotal tbl in reptotals)
            {
                if (tbl.Report_TableJoins != null)
                {
                    hasvalue = true;
                    foreach (string tb in tbl.Report_TableJoins.ToString().Split('-'))
                    {
                        TableJoin t1 = new TableJoin();
                        t1.Commoncolumn = tb.Split('@', '#', ',')[0];
                        t1.Operation = tb.Split('@', '#', ',')[1];
                        t1.table1 = Convert.ToInt32(tb.Split('@', '#', ',')[2]);
                        t1.table2 = Convert.ToInt32(tb.Split('@', '#', ',')[3]);
                        operations.Add(t1);
                    }
                }
            }

            return operations;
        }

        #region IDisposable Members

        public void Dispose()
        {


        }

        #endregion
    }
}
