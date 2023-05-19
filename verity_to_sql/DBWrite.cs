using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Automation;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Internal.ApiImplementation;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Autodesk.Navisworks.Gui.Roamer.CommandLineConfig;

namespace verity_to_sql
{
    public class DBWrite
    {

        ////////////////////////////////////////////////////////////////////////////////////////////
        /////check status in database and if marked installed skip that line.
        public static string GetItemDBStatus(string itemID, SqlConnection inputConn, int existsStatus)
        {
            if (existsStatus == 1)
            {
                /////get status of item currently in database
                string sqlStringstatus = "SELECT verityinstallstatus FROM veritydata WHERE (guid = \'" + itemID + "\')";
                SqlCommand sqlGetStatus = new SqlCommand(sqlStringstatus, inputConn);
                var databaseStatus = sqlGetStatus.ExecuteScalar();
                string databaseStatusString = databaseStatus.ToString();
                return databaseStatusString;
            }
            if (existsStatus > 1)
            {
                MessageBox.Show("GUID was found more than once", "sql integrity issue");
                return "fail";
            }
            return "item not in database";
        }
        ///////////////////////////////////////////////////////////////////////////////////////////

        public static string DBFill(string connString, DataTable data_input)
        {
            try
            {
                SqlConnection wConn = new SqlConnection(connString);
                wConn.Open();

                try
                {
                    DateTime dateTime = DateTime.Now;
                    string cleandate = dateTime.ToString("yyyy-MM-dd");
                    //MessageBox.Show(cleandate, "date");

                    int skippedrows = 0;
                    int badID = 0;

                    DataTable dtWrite = new DataTable();
                    dtWrite.Columns.Add("dtguid");
                    dtWrite.Columns.Add("dtverityinstallstatus");
                    dtWrite.Columns.Add("dtveritynotes");
                    dtWrite.Columns.Add("dtverityguid");
                    dtWrite.Columns.Add("dtdate");

                    DataTable dtUpdate = new DataTable();
                    dtUpdate.Columns.Add("dtguid");
                    dtUpdate.Columns.Add("dtverityinstallstatus");
                    dtUpdate.Columns.Add("dtveritynotes");
                    dtUpdate.Columns.Add("dtverityguid");
                    dtUpdate.Columns.Add("dtdate");

                    foreach (DataRow row in data_input.Rows)
                    {
                        string rowID = row["NavisworksGuid"].ToString();
                        string rowNotes = row["notes"].ToString();
                        string rowStatus = row["Installation Status"].ToString();
                        string rowVID = row["GUID"].ToString();
                        //MessageBox.Show(rowID, "NavisGUID");

                        string badNavisGuid = "00000000-0000-0000-0000-000000000000";

                        if (rowID == badNavisGuid)
                        {
                            badID++;
                        }

                        /////sql statment to check if item exists in database
                        string sqlString = "SELECT COUNT(*) FROM veritydata WHERE (guid = \'" + rowID + "\')";
                        //MessageBox.Show(sqlString, "sql String");
                        SqlCommand check_GUID = new SqlCommand(sqlString, wConn);
                        //check_GUID.Parameters.AddWithValue("@Navisguid", rowID);

                        int GUID_exists = (int)check_GUID.ExecuteScalar();

                        string databaseStatusString = GetItemDBStatus(rowID, wConn, GUID_exists);


                        /////split data in new table and update table
                        if (databaseStatusString.ToLower() != "installed" && rowID != badNavisGuid)
                        {
                            if (GUID_exists > 0)
                            {
                                /////update existing item in database
                                if (rowStatus.ToLower() == "installed")
                                {
                                    object[] newRow = { rowID, rowStatus, rowNotes, rowVID, cleandate };
                                    dtUpdate.Rows.Add(newRow);
                                }
                                else
                                {
                                    object[] newRow = { rowID, rowStatus, rowNotes, rowVID };
                                    dtUpdate.Rows.Add(newRow);
                                }
                            }
                            else
                            {
                                /////write new items to datatable then to database
                                if (rowStatus.ToLower() == "installed")
                                {
                                    object[] newRow = { rowID, rowStatus, rowNotes, rowVID, cleandate };
                                    dtWrite.Rows.Add(newRow);
                                }
                                else
                                {
                                    object[] newRow = { rowID, rowStatus, rowNotes, rowVID };
                                    dtWrite.Rows.Add(newRow);
                                }
                            }
                        }
                        else
                        {
                            skippedrows++;
                        }
                    }////end foreach statement

                    try
                    {
                        int tableRowCount = dtWrite.Rows.Count;
                        if (tableRowCount > 0)
                        {
                            using (SqlBulkCopy bulkCopyTable = new SqlBulkCopy(connString))
                            {

                                bulkCopyTable.DestinationTableName = "dbo.veritydata";

                                /////map col names first datatable then server
                                SqlBulkCopyColumnMapping navisGuid = new SqlBulkCopyColumnMapping("dtguid", "guid");
                                bulkCopyTable.ColumnMappings.Add(navisGuid);
                                SqlBulkCopyColumnMapping vnotes = new SqlBulkCopyColumnMapping("dtveritynotes", "veritynotes");
                                bulkCopyTable.ColumnMappings.Add(vnotes);
                                SqlBulkCopyColumnMapping vstatus = new SqlBulkCopyColumnMapping("dtverityinstallstatus", "verityinstallstatus");
                                bulkCopyTable.ColumnMappings.Add(vstatus);
                                SqlBulkCopyColumnMapping vguid = new SqlBulkCopyColumnMapping("dtverityguid", "verityguid");
                                bulkCopyTable.ColumnMappings.Add(vguid);
                                SqlBulkCopyColumnMapping dtdate = new SqlBulkCopyColumnMapping("dtdate", "date");
                                bulkCopyTable.ColumnMappings.Add(dtdate);

                                try
                                {
                                    bulkCopyTable.WriteToServer(dtWrite);
                                    //wConn.Close();
                                }
                                catch (Exception ex2)
                                {
                                    wConn.Close();
                                    MessageBox.Show(ex2.Message, "Error writing New data table to server");
                                    string itemResult = "Writing Failure";
                                    return itemResult;
                                }
                            }///end using bulk copy
                        }
                        else
                        {

                            //wConn.Close();
                            MessageBox.Show("The new data datatable was empty. Nothing new was written to the database.", "Empty Table");

                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.ToString(), "dtwrite table error");
                    }

                    /////create temp table is sql, write update table to it with bulk copy then join with main table

                    try
                    {
                        SqlConnection updateConn = new SqlConnection(connString);
                        updateConn.Open();
                        string tempTable = "CREATE TABLE #veritytemp(guid varchar(50), veritynotes varchar(50), verityinstallstatus varchar(50), verityguid varchar(50), date varchar(50))";
                        SqlCommand sqlUpdate = new SqlCommand(tempTable, updateConn);

                        sqlUpdate.ExecuteNonQuery();

                        int tableRowCount2 = dtUpdate.Rows.Count;
                        if (tableRowCount2 > 0)
                        {
                            using (SqlBulkCopy bulkCopyTable = new SqlBulkCopy(updateConn))
                            {

                                bulkCopyTable.DestinationTableName = "dbo.#veritytemp";

                                /////map col names first datatable then server
                                SqlBulkCopyColumnMapping navisGuid = new SqlBulkCopyColumnMapping("dtguid", "guid");
                                bulkCopyTable.ColumnMappings.Add(navisGuid);
                                SqlBulkCopyColumnMapping vnotes = new SqlBulkCopyColumnMapping("dtveritynotes", "veritynotes");
                                bulkCopyTable.ColumnMappings.Add(vnotes);
                                SqlBulkCopyColumnMapping vstatus = new SqlBulkCopyColumnMapping("dtverityinstallstatus", "verityinstallstatus");
                                bulkCopyTable.ColumnMappings.Add(vstatus);
                                SqlBulkCopyColumnMapping vguid = new SqlBulkCopyColumnMapping("dtverityguid", "verityguid");
                                bulkCopyTable.ColumnMappings.Add(vguid);
                                SqlBulkCopyColumnMapping dtdate = new SqlBulkCopyColumnMapping("dtdate", "date");
                                bulkCopyTable.ColumnMappings.Add(dtdate);

                                try
                                {
                                    bulkCopyTable.WriteToServer(dtUpdate);

                                    string sqlJoinString = "UPDATE veritydata SET veritydata.veritynotes=#veritytemp.veritynotes, veritydata.verityinstallstatus=#veritytemp.verityinstallstatus, veritydata.verityguid=#veritytemp.verityguid, veritydata.date=#veritytemp.date FROM veritydata INNER JOIN #veritytemp ON veritydata.guid=#veritytemp.guid;";
                                    SqlCommand sqlJoinCom = new SqlCommand(sqlJoinString, updateConn);
                                    sqlJoinCom.ExecuteNonQuery();
                                    //wConn.Close();
                                }
                                catch (Exception ex2)
                                {
                                    wConn.Close();
                                    MessageBox.Show(ex2.Message, "Error writing Update data table to server");
                                    string itemResult = "Writing Failure";
                                    return itemResult;
                                }
                            }///end using bulk copy
                        }
                    }
                    catch (Exception e) 
                    {
                        MessageBox.Show(e.ToString(), "dtupdate table error");
                    }
                    
                    MessageBox.Show(skippedrows.ToString() + " rows not added", "Rows already set to installed.");
                    MessageBox.Show(badID.ToString() + " items with bad Navis GUIDs", "Bad GUIDs");
                }
                catch (Exception ex)
                {
                    wConn.Close();
                    MessageBox.Show(ex.Message, "error checking database");
                    string itemResult2 = "Connection Failed";
                    return itemResult2;
                }
                wConn.Close();
            }   
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error connecting to database.");
                string itemResult2 = "Connection Failed";
                return itemResult2;
            }

            return "Done";
        }
    }

}
