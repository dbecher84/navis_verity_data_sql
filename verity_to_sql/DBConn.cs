using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Automation;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Internal.ApiImplementation;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.AxHost;

namespace verity_to_sql
{
    public class DBConn
    {
        public static string SQL_test(string connString)
        {
            try
            {
                SqlConnection conn = new SqlConnection(connString);
                conn.Open();

                List<string> Table_list = new List<string>();
                using (SqlCommand com = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", conn))
                {
                    using (SqlDataReader reader = com.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Table_list.Add((string)reader["TABLE_NAME"]);
                            //MessageBox.Show((string)reader["TABLE_NAME"], "Table");
                        }
                    }
                }
                conn.Close();
                foreach (string sTable in Table_list)
                {
                    if (Table_list.Contains(sTable.ToLower()))// == "veritydata")
                    {
                        return "Success";
                    }
                    else
                    {
                        MessageBox.Show("Error connecting to datbase. Check network/VPN Connection. Table found " + sTable, "Error");
                        return "Fail";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error connecting to datbase. Check network/VPN Connection.", ex.Message);
                return "Fail";
            }
            return "this return is outside the try statement";
        }
    }

}
