using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Common
{
    public static class ComArguments
    {
        /// <summary>
        /// Given a list of parameters to invoke on a COM object, ensure any null values are converted to Reflection.Missing and any ComWrapper types pass the InternalObject over the boundry.
        /// </summary>
        /// <param name="parms">The parameters to parse.</param>
        /// <returns>An equivalent array with nulls converted to Reflection.Missing.</returns>
        public static object[] Prepare(params object[] parms)
        {
            var result = new object[parms.Length];
            for (int i = 0; i < parms.Length; i++)
                result[i] = ((parms[i] is ComWrapper ? ((ComWrapper)parms[i]).InternalObject : parms[i]) ?? System.Reflection.Missing.Value);

            return result;
        }
    }
}
