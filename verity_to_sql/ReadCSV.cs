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
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Collections.Specialized;
using System.Net.Http;
using System.Data.Common;
using Microsoft.VisualBasic.FileIO;

namespace verity_to_sql
{
    public class ReadCSV
    {
        public static DataTable csvReader(string file_path_name)
        {
            //List<string> csvHeaderList = new List<string>(new string[] { "NavisworksGuid", "Guid", "Installation Status", "Notes" });

            DataTable dt = new DataTable();

            /////Headers in verity csv export used here - Installation Status, Notes, Guid, NavisworksGuid

            TextFieldParser myParser = new TextFieldParser(file_path_name);
            myParser.HasFieldsEnclosedInQuotes = true;
            myParser.SetDelimiters(",");

            string[] headers = myParser.ReadFields();

            foreach (string header in headers)
            {
                dt.Columns.Add(header);
            }

            while (!myParser.EndOfData)
            {
                string[] fields = myParser.ReadFields();

                DataRow dr = dt.NewRow();
                for (int i = 0; i < headers.Length; i++)
                {
                    dr[i] = fields[i];
                }
                dt.Rows.Add(dr);
            }
            myParser.Close();

            return dt;
        }
    }

}
