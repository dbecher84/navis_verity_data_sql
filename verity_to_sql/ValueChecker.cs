using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace verity_to_sql
{
    internal class ValueChecker
    {
        public static string navisValueCheck(DataProperty inputItem)
        {
            if (inputItem != null)
            {
                if (inputItem.Value != null)
                {
                    return inputItem.Value.ToDisplayString();
                }
                else
                {
                    return "Not Available";
                }
            }
            else
            {
                return "fail";
            }
        }
    }
}
