using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Data;
namespace TEMAutomatedReports
{
    class Tableoperations
    {
        public static ArrayList DataSetToArrayList(int ColumnIndex, DataTable dataTable)
        {
            ArrayList output = new ArrayList();

            foreach (DataRow row in dataTable.Rows)
                output.Add(row[ColumnIndex]);

            return output;
        }
        // This operation will jooin the tables.....
        public static DataTable myJoinMethod(DataTable LeftTable, DataTable RightTable,
                String LeftPrimaryColumn, String RightPrimaryColumn)
        {
            //first create the datatable columns 
            DataSet mydataSet = new DataSet();
            mydataSet.Tables.Add("  ");
            DataTable myDataTable = mydataSet.Tables[0];

            //add left table columns 
            DataColumn[] dcLeftTableColumns = new DataColumn[LeftTable.Columns.Count];
            LeftTable.Columns.CopyTo(dcLeftTableColumns, 0);

            foreach (DataColumn LeftTableColumn in dcLeftTableColumns)
            {
                if (!myDataTable.Columns.Contains(LeftTableColumn.ToString()))
                    myDataTable.Columns.Add(LeftTableColumn.ToString(),LeftTableColumn.DataType);
            }

            //now add right table columns 
            DataColumn[] dcRightTableColumns = new DataColumn[RightTable.Columns.Count];
            RightTable.Columns.CopyTo(dcRightTableColumns, 0);

            foreach (DataColumn RightTableColumn in dcRightTableColumns)
            {
                if (!myDataTable.Columns.Contains(RightTableColumn.ToString()))
                {
                    if (RightTableColumn.ToString() != RightPrimaryColumn)
                        myDataTable.Columns.Add(RightTableColumn.ToString(),RightTableColumn.DataType);
                }
            }

            //add left-table data to mytable 
            foreach (DataRow LeftTableDataRows in LeftTable.Rows)
            {
                myDataTable.ImportRow(LeftTableDataRows);
            }

            ArrayList var = new ArrayList(); //this variable holds the id's which have joined 

            ArrayList LeftTableIDs = new ArrayList();
            LeftTableIDs = DataSetToArrayList(0, LeftTable);

            //import righttable which having not equal Id's with lefttable 
            foreach (DataRow rightTableDataRows in RightTable.Rows)
            {
                if (LeftTableIDs.Contains(rightTableDataRows[0]))
                {
                    string wherecondition = "[" + myDataTable.Columns[0].ColumnName + "]='"
                            + rightTableDataRows[0].ToString() + "'";
                    DataRow[] dr = myDataTable.Select(wherecondition);
                    int iIndex = myDataTable.Rows.IndexOf(dr[0]);

                    foreach (DataColumn dc in RightTable.Columns)
                    {
                        if (dc.Ordinal != 0)
                            myDataTable.Rows[iIndex][dc.ColumnName.ToString().Trim()] =
                    rightTableDataRows[dc.ColumnName.ToString().Trim() ?? "0"].ToString();
                    }
                }
                else
                {
                    int count = myDataTable.Rows.Count;
                    DataRow row = myDataTable.NewRow();
                    row[0] = rightTableDataRows[0].ToString();
                    myDataTable.Rows.Add(row);
                    foreach (DataColumn dc in RightTable.Columns)
                    {
                        if (dc.Ordinal != 0)
                            myDataTable.Rows[count][dc.ColumnName.ToString().Trim()] =
                    rightTableDataRows[dc.ColumnName.ToString().Trim() ?? "0"].ToString();
                    }
                }
            }

            foreach (DataRow row in myDataTable.Rows)
            {
                for (int i = 0; i < myDataTable.Columns.Count; i++)
                {
                    if (row[i].ToString() == string.Empty)
                    {
                        row[i] = 0;
                    }

                }

            }
            return myDataTable;
        }

        // this method will determine what operation need to be done on a table
        public static DataSet Combinetables(DataSet ds, string name, string section)
        {
            bool hasvalue;
            List<TableJoin> tJ = Datamethods.Table_Joins(name, section, out hasvalue);
           
            DataSet newdata = new DataSet();
            try
            {
                if (hasvalue == true)
                {
                    foreach (TableJoin table in tJ)
                    {

                        switch (table.Operation)
                        {
                            case "Join":
                                DataTable one = Tableoperations.myJoinMethod(ds.Tables[table.table1], ds.Tables[table.table2], table.Commoncolumn, table.Commoncolumn);
                                DataTable gottam = one.Copy();
                                newdata.Tables.Add(gottam);



                                break;
                            case "Merge":
                                DataTable two = new DataTable();
                                two.Merge(ds.Tables[table.table1], false, MissingSchemaAction.Add);
                                two.Merge(ds.Tables[table.table2], false, MissingSchemaAction.Add);

                                newdata.Tables.Add(two);
                                break;
                            case "NO":
                                DataTable three = ds.Tables[table.table1];
                                DataTable four = three.Copy();
                                newdata.Tables.Add(four);
                                break;
                        }


                    }
                }
                else { newdata = ds; }
            }
            catch { }
            return newdata;
        
        }


    }
}
