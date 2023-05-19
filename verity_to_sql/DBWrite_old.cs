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

namespace verity_to_sql_old
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

                        //MessageBox.Show(databaseStatusString, "database status");

                        if (databaseStatusString.ToLower() != "installed" && rowID != badNavisGuid)
                        {
                            ///sql statement to insert new row
                            string sqlInsertNoDate = "INSERT INTO veritydata (guid, veritynotes, verityinstallstatus, verityguid) VALUES (\'" + rowID + "\',\'" + rowNotes + "\',\'" + rowStatus + "\',\'" + rowVID + "\')";
                            string sqlInsertDate = "INSERT INTO veritydata (guid, veritynotes, verityinstallstatus, verityguid, date) VALUES (\'" + rowID + "\',\'" + rowNotes + "\',\'" + rowStatus + "\',\'" + rowVID + "\',\'" + cleandate + "\')";

                            SqlCommand sqlWriteNoDate = new SqlCommand(sqlInsertNoDate, wConn);
                            SqlCommand sqlWriteDate = new SqlCommand(sqlInsertDate, wConn);

                            /////sql statement to update row
                            string sqlUpdateNoDate = "UPDATE veritydata SET veritynotes=\'" + rowNotes + "\',verityinstallstatus=\'" + rowStatus + "\',verityguid=\'" + rowVID + "\' WHERE guid = \'" + rowID + "\'";
                            string sqlUpdateDate = "UPDATE veritydata SET veritynotes=\'" + rowNotes + "\',verityinstallstatus=\'" + rowStatus + "\',verityguid=\'" + rowVID + "\',date=\'" + cleandate + "\' WHERE guid = \'" + rowID + "\'";
                            //MessageBox.Show(sqlUpdateNoDate, "sql Update String");

                            SqlCommand sqlWriteUpdateNoDate = new SqlCommand(sqlUpdateNoDate, wConn);
                            SqlCommand sqlWriteUpdateDate = new SqlCommand(sqlUpdateDate, wConn);

                            if (GUID_exists > 0)
                            {
                                //////update existing item in database
                                if (rowStatus.ToLower() == "installed")
                                {
                                    sqlWriteUpdateDate.ExecuteNonQuery();
                                }
                                else
                                {
                                    sqlWriteUpdateNoDate.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                /////write new item to database
                                if (rowStatus.ToLower() == "installed")
                                {
                                    sqlWriteDate.ExecuteNonQuery();
                                }
                                else
                                {
                                    sqlWriteNoDate.ExecuteNonQuery();
                                }
                            }
                        }
                        else
                        {
                            skippedrows ++;
                        }
                    }
                    MessageBox.Show(skippedrows.ToString() + " rows not changed", "Rows already set to install.");
                    MessageBox.Show(badID.ToString() + " items with bad Navis GUIDs", "Bad GUIDs");
                }
                catch (Exception ex)
                {
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
