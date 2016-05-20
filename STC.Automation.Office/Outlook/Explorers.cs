using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Outlook.Enums;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Contains a set of Explorer objects representing all explorers.
    /// </summary>
    [WrapsCOM("Outlook.Explorers", "0006300A-0000-0000-C000-000000000046")]
    public class Explorers : ComWrapper
    {
        internal Explorers(object explorersObj)
            : base(explorersObj)
        {
        }

        /// <summary>
        /// Creates a new recipient in the Recipients collection.
        /// </summary>
        /// <param name="folder">The Variant object to display in the explorer window when it is created.</param>
        /// <param name="displayMode">The display mode of the folder.</param>
        /// <returns>An Explorer object that represents a new instance of the window.</returns>
        public Explorer Add(Folder folder, FolderDisplayMode displayMode)
        {
            return new Explorer(InternalObject.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { folder.InternalObject, displayMode }));
        }
        
        /// <summary>
        /// Index the Explorers collection to get an explorer.
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public Explorer this[int key]
        {
            get
            {
                try
                {
                    return new Explorer(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { key }));
                }
                catch (Exception ex)
                {
                    throw new IndexOutOfRangeException(String.Concat("Could not find explorer '", key, "'."), ex);
                }
            }
        }
        
        /// <summary>
        /// Returns an integer value that represents the number of objects in the collection.
        /// </summary>
        public int Count
        {
            get
            {
                return (int)InternalObject.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }
        }
    }
}
