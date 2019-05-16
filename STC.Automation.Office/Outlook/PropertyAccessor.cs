using STC.Automation.Office.Attributes;
using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Provides the ability to create, get, set, and delete properties on objects.
    /// </summary>
    [WrapsCOM("Outlook.PropertyAccessor", "0006302D-0000-0000-C000-000000000046")]
    public class PropertyAccessor : ComWrapper
    {
        internal PropertyAccessor(object attachmentObj)
            : base(attachmentObj)
        {
        }

        /// <summary>
        /// Sets the property specified by SchemaName to the value specified by Value .
        /// </summary>
        /// <param name="schemaName">The name of a property whose value is to be set as specified by the Value parameter. The property is referenced by namespace. For more information, see https://docs.microsoft.com/en-us/office/vba/outlook/how-to/navigation/referencing-properties-by-namespace </param>
        /// <param name="value">The value that is to be set for the property specified by the SchemaName parameter.</param>
        public void SetProperty(string schemaName, object value)
        {
            InternalObject.GetType().InvokeMember("SetProperty", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { schemaName, value });
        }
    }
}
