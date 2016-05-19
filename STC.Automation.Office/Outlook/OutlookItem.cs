using STC.Automation.Office.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Parent class for MailItem to aid in CreateItem and traversal of items.  Does not correspond to any object in the Outlook object model.
    /// </summary>
    [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
    public abstract class OutlookItem : ComWrapper
    {
        internal OutlookItem(object itemObj)
            : base(itemObj)
        {
        }

        /// <summary>
        /// Factory method to create and return the correct OutlookItem object (MailItem) for the given COM object.
        /// </summary>
        /// <param name="comObj"></param>
        /// <returns></returns>
        internal static OutlookItem ResolveType(object comObj)
        {
            if (ComWrapper.SupportsInterface(comObj, ComWrapper.GetMustSupport(typeof(MailItem))))
                return new MailItem(comObj);

            throw new NotImplementedException("Unknown object when attempting to resolve Outlook item type.");
        }
    }
}
