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
    [AddInPluginAttribute(AddInLocation.AddIn, Icon = "C:\\Program Files\\Autodesk\\Navisworks Manage 2023\\Plugins\\verity_to_sql\\resources\\16x16_verity_export_img.bmp",
        LargeIcon = "C:\\Program Files\\Autodesk\\Navisworks Manage 2023\\Plugins\\verity_to_sql\\resources\\32x32_verity_export_img.bmp")]

    public class SqlWrite : AddInPlugin
    {
        static string dbconn = "Data Source = USBLB1DB002\\APP05;Initial Catalog=VerityData;Integrated Security=true";

        public override int Execute(params string[] parameters)
        {
            //Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;

            bool proceed = true;

            string dbTestResult = DBConn.SQL_test(dbconn);

            ////////Form to ask user if data is from verity csv or navisworks////////////////////////////////
            //var sourceInfo = new DataSource();
            //sourceInfo.ShowDialog();

            //string useCSV = sourceInfo.usecsvdata;
            //string useNavis = sourceInfo.usenavisdata;  

            //MessageBox.Show("Use csv? = " + useCSV  + " Use Navis? = " + useNavis, "button check");
            ///////////////////////////////////////////////////////////////////////////////////////////////
            string inputCsvFile = null;

            if (dbTestResult == "Success")
            {
                //////Form to ask user if data is from verity csv or navisworks////////////////////////////////
                var sourceInfo = new DataSource();
                sourceInfo.ShowDialog();

                string useCSV = sourceInfo.usecsvdata;
                string useNavis = sourceInfo.usenavisdata;

                ////////////////////////////////////////////////////////////////////////////////////////////                
                if (useCSV == "true")
                {
                    inputCsvFile = GetCSVInfo.csvFileSelector();

                    if (string.IsNullOrEmpty(inputCsvFile))
                    {
                        MessageBox.Show("No csv file was selected.", "File Error");
                        proceed = false;
                    }
                    if (proceed == true && inputCsvFile.ToString().Substring(inputCsvFile.Length - 3) != "csv")
                    {
                        //MessageBox.Show(inputCsvFile.ToString().Substring(inputCsvFile.Length - 3), "test");
                        MessageBox.Show("The file selected was not a csv file.", "File Error");
                        proceed = false;
                    }
                }
                ////////////////////////////////////////////////////////////////////////////////////////////

                ////////////////////////////////////////////////////////////////////////////////////////////
                //////project number and model name from manual entry
                //string manualProjectNumber = null;
                //string manualModelName = null;

                if (proceed == true)
                {
                    var pInfo = new ManualInfo();
                    pInfo.ShowDialog();

                    //manualProjectNumber = ManualInfo.ProjectNumForm;
                    //manualModelName = ManualInfo.ModelNameForm;
                    //MessageBox.Show("NUM= " + manualProjectNumber + " ,Name= " + manualModelName, "Model Name Error");

                    //////project number model name auto collected from file path
                    //string projectNumber = GetProjectInfo.GetProjectNum();
                    //string modelName = GetProjectInfo.GetNavisName();

                    if (string.IsNullOrEmpty(ManualInfo.ProjectNumForm) || string.IsNullOrEmpty(ManualInfo.ModelNameForm))
                    {
                        MessageBox.Show("Either the project number, the model name, or both were not entered", "Missing Information");
                        proceed = false;
                    }
                    /////character count is limited by the windows form settings
                    //if (manualModelName.Count() > 50)
                    //{
                    //    MessageBox.Show("The model name that was entered is too long.", "Model Name Error");
                    //    proceed = false;
                    //}
                }
                ////////////////////////////////////////////////////////////////////////////////////////////

                if (proceed == true)
                {
                    //////datatable from csv file
                    if (useCSV == "true")
                    {
                        DataTable csvTable = ReadCSV.csvReader(inputCsvFile);

                        DBWrite.DBFill(dbconn, csvTable, ManualInfo.ProjectNumForm, ManualInfo.ModelNameForm, ManualInfo.BuildingNameForm);
                    }
                    if (useNavis == "true")
                    {
                        DataTable navisTable = ReadNavisData.NavisReader();

                        DBWrite.DBFill(dbconn, navisTable, ManualInfo.ProjectNumForm, ManualInfo.ModelNameForm, ManualInfo.BuildingNameForm);

                    }
                }
                else
                {
                    MessageBox.Show("Program terminated due to missing/bad information.", "Missing Information");
                }
            }
            else
            {
                MessageBox.Show("Program terminated due to bad connection to database.", "Error");
            }
            return 0;
        }
    }

}
