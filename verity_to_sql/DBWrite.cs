using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Automation;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Internal.ApiImplementation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
                string sqlStringstatus = "SELECT verityinstallstatus FROM veritydata WHERE (id = \'" + itemID + "\')";
                SqlCommand sqlGetStatus = new SqlCommand(sqlStringstatus, inputConn);
                var databaseStatus = sqlGetStatus.ExecuteScalar();
                string databaseStatusString = databaseStatus.ToString();
                return databaseStatusString;
            }
            if (existsStatus > 1)
            {
                MessageBox.Show("Item id ID was found more than once", "sql integrity issue");
                return "fail";
            }
            return "item not in database";
        }
        ///////////////////////////////////////////////////////////////////////////////////////////

        public static string DBFill(string connString, DataTable data_input, string input_projectnumber, string input_modelname, string input_bldgname)
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
                    dtWrite.Columns.Add("dtnavisguid");
                    dtWrite.Columns.Add("dtnverityinstallstatus");
                    dtWrite.Columns.Add("dtnveritynotes");
                    dtWrite.Columns.Add("dtnverityguid");
                    dtWrite.Columns.Add("dtndate"); 
                    dtWrite.Columns.Add("dtnprojectnum");
                    dtWrite.Columns.Add("dtnmodelname");
                    dtWrite.Columns.Add("dtnbldgname");
                    dtWrite.Columns.Add("dtnfuture");

                    DataTable dtUpdate = new DataTable();
                    dtUpdate.Columns.Add("dtnavisguid");
                    dtUpdate.Columns.Add("dtverityinstallstatus");
                    dtUpdate.Columns.Add("dtveritynotes");
                    dtUpdate.Columns.Add("dtverityguid");
                    dtUpdate.Columns.Add("dtdate");
                    dtUpdate.Columns.Add("dtprojectnum");
                    dtUpdate.Columns.Add("dtmodelname");
                    dtUpdate.Columns.Add("dtbldgname");
                    dtUpdate.Columns.Add("dtfuture");

                    foreach (DataRow row in data_input.Rows)
                    {
                        string rowNID = row["NavisworksGuid"].ToString();
                        string rowNotes = row["notes"].ToString();
                        string rowStatus = row["Installation Status"].ToString();
                        string rowVID = row["GUID"].ToString();
                        string rowNVID = "future use";
                        //MessageBox.Show(rowID, "NavisGUID");

                        string badNavisGuid = "00000000-0000-0000-0000-000000000000";
                        //string blankNavidGuid = "";

                        if (rowNID == badNavisGuid)
                        {
                            badID++;
                        }

                        /////sql statment to check if item exists in database
                        string sqlString = "SELECT COUNT(*) FROM veritydata WHERE (id = \'" + rowNID + "\')";
                        //MessageBox.Show(sqlString, "sql String");
                        SqlCommand check_GUID = new SqlCommand(sqlString, wConn);
                        //check_GUID.Parameters.AddWithValue("@Navisguid", rowID);

                        int GUID_exists = (int)check_GUID.ExecuteScalar();

                        string databaseStatusString = GetItemDBStatus(rowNID, wConn, GUID_exists);


                        /////split data in new table and update table
                        if (databaseStatusString.ToLower() != "installed" && rowNID != badNavisGuid && string.IsNullOrEmpty(rowNID) == false)
                        {
                            if (GUID_exists > 0)
                            {
                                /////update existing item in database
                                if (rowStatus.ToLower() == "installed")
                                {
                                    object[] newRow = { rowNID, rowStatus, rowNotes, rowVID, cleandate, input_projectnumber, input_modelname, input_bldgname, rowNVID };
                                    dtUpdate.Rows.Add(newRow);
                                }
                                else
                                {
                                    object[] newRow = { rowNID, rowStatus, rowNotes, rowVID, "no install date", input_projectnumber, input_modelname, input_bldgname, rowNVID };
                                    dtUpdate.Rows.Add(newRow);
                                }
                            }
                            else
                            {
                                /////write new items to datatable then to database
                                if (rowStatus.ToLower() == "installed")
                                {
                                    object[] newRow = { rowNID, rowStatus, rowNotes, rowVID, cleandate, input_projectnumber, input_modelname, input_bldgname, rowNVID };
                                    dtWrite.Rows.Add(newRow);
                                }
                                else
                                {
                                    object[] newRow = { rowNID, rowStatus, rowNotes, rowVID, "no install date", input_projectnumber, input_modelname, input_bldgname, rowNVID };
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
                            using (SqlBulkCopy bulkCopyTableNew = new SqlBulkCopy(connString))
                            {
                                bulkCopyTableNew.DestinationTableName = "dbo.veritydata";

                                /////map col names first datatable then server
                                SqlBulkCopyColumnMapping id = new SqlBulkCopyColumnMapping("dtnavisguid", "id");
                                bulkCopyTableNew.ColumnMappings.Add(id);
                                SqlBulkCopyColumnMapping navisGuid = new SqlBulkCopyColumnMapping("dtnfuture", "navisguid");
                                bulkCopyTableNew.ColumnMappings.Add(navisGuid);
                                SqlBulkCopyColumnMapping vnotes = new SqlBulkCopyColumnMapping("dtnveritynotes", "veritynotes");
                                bulkCopyTableNew.ColumnMappings.Add(vnotes);
                                SqlBulkCopyColumnMapping vstatus = new SqlBulkCopyColumnMapping("dtnverityinstallstatus", "verityinstallstatus");
                                bulkCopyTableNew.ColumnMappings.Add(vstatus);
                                SqlBulkCopyColumnMapping vguid = new SqlBulkCopyColumnMapping("dtnverityguid", "verityguid");
                                bulkCopyTableNew.ColumnMappings.Add(vguid);
                                SqlBulkCopyColumnMapping dtdate = new SqlBulkCopyColumnMapping("dtndate", "date");
                                bulkCopyTableNew.ColumnMappings.Add(dtdate);
                                SqlBulkCopyColumnMapping projectnum = new SqlBulkCopyColumnMapping("dtnprojectnum", "projectnumber");
                                bulkCopyTableNew.ColumnMappings.Add(projectnum);
                                SqlBulkCopyColumnMapping model = new SqlBulkCopyColumnMapping("dtnmodelname", "modelname");
                                bulkCopyTableNew.ColumnMappings.Add(model);
                                SqlBulkCopyColumnMapping building = new SqlBulkCopyColumnMapping("dtnbldgname", "buildingname");
                                bulkCopyTableNew.ColumnMappings.Add(building);

                                try
                                {
                                    bulkCopyTableNew.WriteToServer(dtWrite);
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
                        string tempTable = "CREATE TABLE #veritytemp(id varchar(255), veritynotes varchar(50), verityinstallstatus varchar(50), verityguid varchar(50), date varchar(50), projectnumber varchar(50), modelname varchar(50), buildingname varchar(50), navisguid varchar(50))";
                        SqlCommand sqlUpdate = new SqlCommand(tempTable, updateConn);
                        //MessageBox.Show("Temp table created", "success");

                        sqlUpdate.ExecuteNonQuery();

                        int tableRowCount2 = dtUpdate.Rows.Count;
                        if (tableRowCount2 > 0)
                        {
                            using (SqlBulkCopy bulkCopyTableUpdate = new SqlBulkCopy(updateConn))
                            {

                                bulkCopyTableUpdate.DestinationTableName = "dbo.#veritytemp";

                                /////map col names first datatable then server
                                SqlBulkCopyColumnMapping id = new SqlBulkCopyColumnMapping("dtnavisguid", "id");
                                bulkCopyTableUpdate.ColumnMappings.Add(id);
                                SqlBulkCopyColumnMapping navisGuid = new SqlBulkCopyColumnMapping("dtfuture", "navisguid");
                                bulkCopyTableUpdate.ColumnMappings.Add(navisGuid);
                                SqlBulkCopyColumnMapping vnotes = new SqlBulkCopyColumnMapping("dtveritynotes", "veritynotes");
                                bulkCopyTableUpdate.ColumnMappings.Add(vnotes);
                                SqlBulkCopyColumnMapping vstatus = new SqlBulkCopyColumnMapping("dtverityinstallstatus", "verityinstallstatus");
                                bulkCopyTableUpdate.ColumnMappings.Add(vstatus);
                                SqlBulkCopyColumnMapping vguid = new SqlBulkCopyColumnMapping("dtverityguid", "verityguid");
                                bulkCopyTableUpdate.ColumnMappings.Add(vguid);
                                SqlBulkCopyColumnMapping dtdate = new SqlBulkCopyColumnMapping("dtdate", "date");
                                bulkCopyTableUpdate.ColumnMappings.Add(dtdate);
                                SqlBulkCopyColumnMapping projectnum = new SqlBulkCopyColumnMapping("dtprojectnum", "projectnumber");
                                bulkCopyTableUpdate.ColumnMappings.Add(projectnum);
                                SqlBulkCopyColumnMapping model = new SqlBulkCopyColumnMapping("dtmodelname", "modelname");
                                bulkCopyTableUpdate.ColumnMappings.Add(model);
                                SqlBulkCopyColumnMapping building = new SqlBulkCopyColumnMapping("dtbldgname", "buildingname");
                                bulkCopyTableUpdate.ColumnMappings.Add(building);

                                try
                                {
                                    bulkCopyTableUpdate.WriteToServer(dtUpdate);

                                    string sqlJoinString = "UPDATE veritydata SET veritydata.veritynotes=#veritytemp.veritynotes, veritydata.verityinstallstatus=#veritytemp.verityinstallstatus, veritydata.verityguid=#veritytemp.verityguid, veritydata.date=#veritytemp.date, veritydata.projectnumber=#veritytemp.projectnumber, veritydata.modelname=#veritytemp.modelname, veritydata.buildingname=#veritytemp.buildingname, veritydata.navisguid=#veritytemp.navisguid FROM veritydata INNER JOIN #veritytemp ON veritydata.id=#veritytemp.id;";
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

                    if (skippedrows > 0)
                    {
                        MessageBox.Show(skippedrows.ToString() + " rows not added", "Rows already set to installed.");
                    }
                    if (badID > 0)
                    {
                        MessageBox.Show(badID.ToString() + " items with bad Navis GUIDs", "Bad GUIDs");
                    }
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
