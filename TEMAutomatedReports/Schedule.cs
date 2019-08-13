using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TEMAutomatedReports
{
   public class  Schedule
    {
        private int _intID;
        private int _autointID;
        private string _struserselectedname = "";
        private string _strStoredProcedure = "";
        private string _strReportName = "";
        private DateTime _dtCreatedDate;
        private string _strListofFilters = "";
        private string _strFrequency = "";
        private string _strType = "";
        private string _strEmailAddresses = "";
        private string _strColumns = "";
        private string _strGraphBindings = "";
        private string _chosenNodeIds = "";
        private string _chosentotals = "";
        private int _dropdownid;
        private int _portfolioreportid;
        private bool _reportstatus;
        private int _reporttime;
        public string _graphtype;
        public string _reportingsection;
        public int _intUserid;
        public string _GraphHeaders;

        public int Time
        {
            get { return _reporttime; }

            set { _reporttime = value; }
        }

        public bool Status
        {
            get { return _reportstatus; }

            set { _reportstatus = value; }
        }

        public string Totals
        {
            get { return _chosentotals; }

            set { _chosentotals = value; }
        }
        public string Chosennodelist
        {
            get { return _chosenNodeIds; }

            set { _chosenNodeIds = value; }
        }
        public string Selectedname
        {
            get { return _struserselectedname; }

            set { _struserselectedname = value; }
        }
        public Int32 ID1
        {
            get { return _autointID; }
            set { _autointID = value; }
        }
        public Int32 ID
        {
            get { return _intID; }
            set { _intID = value; }
        }
        public string StoredProcedureName
        {
            get { return _strStoredProcedure; }
            set { _strStoredProcedure = value; }
        }

        public string ReportName
        {
            get { return _strReportName; }
            set { _strReportName = value; }
        }

        public DateTime CreatedDate
        {
            get { return _dtCreatedDate; }
            set { _dtCreatedDate = value; }
        }

        public string ListofFilters
        {
            get { return _strListofFilters; }
            set { _strListofFilters = value; }
        }

        public string Frequency
        {
            get { return _strFrequency; }
            set { _strFrequency = value; }
        }

        public string Type
        {
            get { return _strType; }
            set { _strType = value; }
        }

        public string EmailAddresses
        {
            get { return _strEmailAddresses; }
            set { _strEmailAddresses = value; }
        }

        public string Columns
        {
            get { return _strColumns; }
            set { _strColumns = value; }
        }


        public string GraphBindings
        {
            get { return _strGraphBindings; }
            set { _strGraphBindings = value; }
        }

        public int Dropdownid
        {
            get { return _dropdownid; }
            set { _dropdownid = value; }
        }
        public int Portfolioreportid
        {
            get { return _portfolioreportid; }
            set { _portfolioreportid = value; }
        }
        public string GraphType
        {
            get { return _graphtype; }
            set { _graphtype = value; }
        }
        public string ReportingSection
        {
            get { return _reportingsection; }
            set { _reportingsection = value; }
        }
        public int UserId
        {
            get { return _intUserid; }

            set { _intUserid = value; }
        }
        public string GraphHeaders
        {
            get { return _GraphHeaders; }

            set { _GraphHeaders = value; }
        }
    }



   public class ColumnTotals
   {
       public String ColumnName { get; set; }
       public String Friendlyname { get; set; }
       public String Totaltype { get; set; }


   }
   public class GraphHeaders
   {
       public String ColumnName { get; set; }
       public String Headername { get; set; }

   }
   public class TableJoin
   {
       public String Commoncolumn { get; set; }
       public String Operation { get; set; }
       public int table1 { get; set; }
       public int table2 { get; set; }
   }
   public enum TotalTypes
   {
       Avg = 1,
       Count,
       First,
       Last,
       Max,
       Min,
       Sum,
       Percentage,
       Avgduration,
       Totalduration,
       None = -1

   }
   
}
