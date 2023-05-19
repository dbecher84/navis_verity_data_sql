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
using System.Data;
using System.IO;
using System.Data.SqlTypes;

namespace verity_to_sql
{
    //plugin attributes require Name, DeveloperID and optional parameters
    [PluginAttribute("verity_sql", "Derek B", DisplayName = "Export Verity Data to SQL", ToolTip = "Exports Verity Data to SQL.", ExtendedToolTip = "Version 2023.0.0.0")]
    [AddInPluginAttribute(AddInLocation.AddIn, Icon = "",
        LargeIcon = "")]

    public class SqlWrite : AddInPlugin
    {
        static string dbconn = "Data Source = USBLB1DB002\\APP05;Initial Catalog=VerityData;Integrated Security=true";

        public static string csvFileSelector()
        {
            //var fileContent = string.Empty;
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    ////Read the contents of the file into a stream
                    //var fileStream = openFileDialog.OpenFile();

                    //using (StreamReader reader = new StreamReader(fileStream))
                    //{
                    //    fileContent = reader.ReadToEnd();
                    //}
                }
            }
            return filePath;
        }

        public override int Execute(params string[] parameters)
        {
            Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;

            string dbTestResult = DBConn.SQL_test(dbconn);

            if (dbTestResult == "Success")
            {
                string inputCsvFile = csvFileSelector();

               // MessageBox.Show(inputCsvFile, "file selected");

                ///////hard coded file path for csv
                //string csv_file_path = @"C:\Users\dbecher\Documents\GitHub\verity_to_sql\test_csv";
                //string file_name = "ElementProperties_Bonland.csv";
                //string file_name = "ElementProperties_Bonland_full.csv";
                //string filePathName = csv_file_path + "\\" + file_name;

                if (string.IsNullOrEmpty(inputCsvFile))
                {
                    MessageBox.Show("No file was selected.", "File Error");
                }
                if (inputCsvFile.ToString().Substring(inputCsvFile.Length -3) != "csv")
                {
                    //MessageBox.Show(inputCsvFile.ToString().Substring(inputCsvFile.Length - 3), "test");
                    MessageBox.Show("The file selected was not a csv file.", "File Error");
                }
                else
                {
                    DataTable myData = ReadCSV.csvReader(inputCsvFile);

                    DBWrite.DBFill(dbconn, myData);
                }
            }

            return 0;
        }
    }

}
