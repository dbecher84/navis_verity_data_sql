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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace verity_to_sql
{
    /// <summary>
    /// pull project number and navis file name from navisworks model file path
    /// may need to be manual due to the abilitly to run this without a model open
    /// </summary>
    public class GetProjectInfo
    {
        public static string navisFilePath = Autodesk.Navisworks.Api.Application.MainDocument.FileName.ToString();

        public static string GetProjectNum()
        {
            try
            {
                ////Get project number from navis file path
                var regMatchNum = @"([0-9]{5})";
                Match regMatch = Regex.Match(navisFilePath, regMatchNum);

                if (regMatch.Success)
                {
                    return regMatch.Value;
                }
                else
                {
                    MessageBox.Show("No project number was found in the file path.", "No project number" );
                    return "Fail";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error getting Project Number.", ex.Message);
                return "Fail";
            }
        }

        public static string GetNavisName()
        {
            try
            {
                ////Get Navis model name from navis file path
                int startPosition = navisFilePath.LastIndexOf("\\") + 1;
                int endPosition = navisFilePath.LastIndexOf(".");
                string modelName = navisFilePath.Substring(startPosition, endPosition - startPosition);

                if (string.IsNullOrEmpty(modelName))
                {
                    MessageBox.Show("Unable to get model from file path.", "No model name");
                    return "Fail";
                }
                else
                {
                    return modelName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error getting Model Name", ex.Message);
                return "Fail";
            }
        }
    }

}
