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
    public class ReadNavisData
    {
        public static DataTable NavisReader()
        {
            Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;

            DataTable navisDT = new DataTable();
            navisDT.Columns.Add("NavisworksGuid");
            navisDT.Columns.Add("notes");
            navisDT.Columns.Add("Installation Status");
            navisDT.Columns.Add("GUID");
            //navisDT.Columns.Add("dtndate");
            //navisDT.Columns.Add("dtnprojectnum");
            //navisDT.Columns.Add("dtnmodelname");

            string verityInternalTabName = "LcOaPropOverrideCat";
            string verityInternalGuid = "VerityGuid";
            string verityInternalStatus = "VerityClassification";
            string verityInternalNotes = "VerityNotes";

            string navisInternalItemTab = "LcOaNode";
            string navisInternalGuid = "LcOaNodeGuid";

            List<string> test = new List<string>();

            ////////future add selection to program to eliminate the search set in Navisworks///////////////////////////////
            //Search mySearch = new Search();
            //SearchCondition itemGuidtest = SearchCondition.HasPropertyByName(navisInternalItemTab, navisInternalGuid);
            //SearchCondition verityGuidtest = SearchCondition.HasPropertyByName(verityInternalTabName, verityInternalGuid);
            //mySearch.SearchConditions.Add(itemGuidtest);
            //mySearch.SearchConditions.Add(verityGuidtest);

            //mySearch.Selection.SelectAll();
            //mySearch.Locations = SearchLocations.DescendantsAndSelf;

            //ModelItemCollection mySearchItems = mySearch.FindAll(document, true);
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

            foreach (ModelItem item in document.CurrentSelection.SelectedItems)
            {

                PropertyCategory itemCat = item.PropertyCategories.FindCategoryByName(verityInternalTabName);
                DataProperty navisGuid = item.PropertyCategories.FindPropertyByName(navisInternalItemTab, navisInternalGuid);

                if (itemCat != null)
                {
                    DataProperty verityGuid = itemCat.Properties.FindPropertyByName(verityInternalGuid);
                    DataProperty verityStatus = itemCat.Properties.FindPropertyByName(verityInternalStatus);
                    DataProperty verityNotes = itemCat.Properties.FindPropertyByName(verityInternalNotes);

                    var row = navisDT.NewRow();
                    string navisGuidTest = ValueChecker.navisValueCheck(navisGuid);
                    if (navisGuidTest != "fail")
                    {
                        row["NavisworksGuid"] = navisGuid.Value.ToDisplayString();
                    }
                    string verityStatusTest = ValueChecker.navisValueCheck(verityStatus);
                    if (verityStatusTest != "fail")
                    {
                        row["Installation Status"] = verityStatus.Value.ToDisplayString();
                    }
                    string verityNotesTest = ValueChecker.navisValueCheck(verityNotes);
                    if (verityNotesTest != "fail")
                    {
                        row["notes"] = verityNotes.Value.ToDisplayString();
                    }
                    string verityGuidTest = ValueChecker.navisValueCheck(verityGuid);
                    if (verityGuidTest != "fail")
                    {
                        row["GUID"] = verityGuid.Value.ToDisplayString();
                    }
                    //if (verityStatus.Value.ToDisplayString().ToLower() == "installed")
                    //{
                    //    row["dtndate"] = cleandate;
                    //}

                    navisDT.Rows.Add(row);

                    test.Add(verityGuid.Value.ToString());
                }
            }

            //int listLen = test.Count;
            int tableCount = navisDT.Rows.Count;
            MessageBox.Show("Number of item to write/update in database " + tableCount.ToString(), "Number of Items");
            //MessageBox.Show(testdescen.Count.ToString(), "Number of Items");

            return navisDT;
        }
    }

}
