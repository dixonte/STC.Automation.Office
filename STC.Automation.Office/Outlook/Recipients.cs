using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Contains a collection of Recipient objects for an Outlook item.
    /// </summary>
    [WrapsCOM("Outlook.Recipients", "0006303B-0000-0000-C000-000000000046")]
    public class Recipients : ComWrapper
    {
        internal Recipients(object recipientsObj)
            : base(recipientsObj)
        {
        }

        /// <summary>
        /// Creates a new recipient in the Recipients collection.
        /// </summary>
        /// <param name="name">The name of the recipient; it can be a string representing the display name, the alias, or the full SMTP e-mail address of the recipient.</param>
        /// <returns>A Recipient object that represents the new recipient.</returns>
        public Recipient Add(string name)
        {
            return new Recipient(InternalObject.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { name }));
        }

        /// <summary>
        /// Attempts to resolve all the Recipient objects in the Recipients collection against the Address Book.
        /// </summary>
        /// <returns>True if all of the objects were resolved, False if one or more were not.</returns>
        public bool ResolveAll()
        {
            return (bool)InternalObject.GetType().InvokeMember("ResolveAll", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
        }
        
        /// <summary>
        /// Index the Recipients collection to get a recipient
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public Recipient this[int key]
        {
            get
            {
                try
                {
                    return new Recipient(InternalObject.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { key }));
                }
                catch (Exception ex)
                {
                    throw new IndexOutOfRangeException(String.Concat("Could not find recipient '", key, "'."), ex);
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

        /// <summary>
        /// Gets a generic IEnumerator of Recipient objects.
        /// </summary>
        /// <returns>IEnumerator&lt;Recipient&gt;</returns>
        //public IEnumerator<Recipient> GetEnumerator()
        //{
        //    return new ComIEnumeratorWrapper<Recipient>(InternalObject.GetType().InvokeMember("_NewEnum", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
        //}
    }
}
