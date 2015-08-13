using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using System.Reflection;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Parent class for Worksheet and Chart to aid in object traversal through Sheets collection.  Does not correspond to any object in the Excel object model.
    /// </summary>
    [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
    public abstract class Sheet : ComWrapper
    {
        internal Sheet(object worksheetObj)
            : base(worksheetObj)
        {
        }

        /// <summary>
        /// Gets or sets a String value representing the name of the sheet.
        /// </summary>
        public string Name
        {
            get
            {
                return InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null) as string;
            }
            set
            {
                InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Factory method to create and return the correct Sheet object (Worksheet or Chart) for the given COM object.
        /// </summary>
        /// <param name="comObj"></param>
        /// <returns></returns>
        internal static Sheet ResolveType(object comObj)
        {
            if (ComWrapper.SupportsInterface(comObj, ComWrapper.GetMustSupport(typeof(Worksheet))))
                return new Worksheet(comObj);
            else if (ComWrapper.SupportsInterface(comObj, ComWrapper.GetMustSupport(typeof(Chart))))
                return new Chart(comObj);

            throw new NotImplementedException("Unknown object when attempting to resolve sheet type.");
        }
    }
}
