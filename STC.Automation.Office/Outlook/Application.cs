using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Diagnostics;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;

namespace STC.Automation.Office.Outlook
{
    /// <summary>
    /// Wraps the Outlook.Application COM object
    /// </summary>
    [WrapsCOM("Outlook.Application", Application.UUID)]
    public class Application : OfficeApplication
    {
        //public const string UUID = "0006F03A-0000-0000-C000-000000000046"; // < -this is the official UUID but doesn't work on mine
        public const string UUID = "00063001-0000-0000-C000-000000000046";

        private Explorers _explorers;

        /// <summary>
        /// Creates a new instance of Excel for the purposes of automation
        /// </summary>
        public Application()
            : base()
        {
        }

        internal Application(object applicationObj)
            : base(applicationObj)
        {
        }

        /// <summary>
        /// Creates and returns a new Microsoft Outlook item.
        /// </summary>
        /// <remarks>The CreateItem method can only create default Outlook items. To create new items using a custom form, use the Add method on the Items collection.</remarks>
        /// <param name="itemType">The Outlook item type for the new item.</param>
        public OutlookItem CreateItem(Enums.ItemType itemType)
        {
            return OutlookItem.ResolveType(InternalObject.GetType().InvokeMember("CreateItem", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { itemType }));
        }

        /// <summary>
        /// Returns a NameSpace object of the specified type.
        /// </summary>
        /// <param name="type">The type of name space to return</param>
        public NameSpace GetNameSpace(string type)
        {
            return new NameSpace(InternalObject.GetType().InvokeMember("GetNamespace", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { type }));
        }


        /// <summary>
        /// Attempts to attach to an already running Outlook process.
        /// </summary>
        /// <param name="processToAttach">The Process object to which to attach.</param>
        /// <returns>An Application wrapper.</returns>
        public static Application FromProcess(Process processToAttach)
        {
            using (Application application = ComWrapper.FromProcess<Application>(processToAttach, "OUTLOOK"))
            {
                if (application != null)
                {
                    return application;
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets a list of all running Outlook applications from the Running Object Table... in theory. Each instance should be manually disposed.
        /// </summary>
        /// <returns>A list of Excel.Application objects</returns>
        public static IList<Application> GetRunningApplications()
        {
            return Application.FromROT<Application>();
        }


        /// <summary>
        /// Returns the topmost Explorer object on the desktop. This object must be manually disposed.
        /// </summary>
        public Explorer ActiveExplorer
        {
            get
            {
                var obj = InternalObject.GetType().InvokeMember("ActiveExplorer", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, null);
                if (obj != null)
                    return new Explorer(obj);
                return null;
            }
        }

        /// <summary>
        /// Returns an Explorers collection object that contains the Explorer objects representing all open explorers. This object is internally cached and does not require manual disposal.
        /// </summary>
        public Explorers Explorers
        {
            get
            {
                if (_explorers == null)
                {
                    _explorers = new Explorers(InternalObject.GetType().InvokeMember("Explorers", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
                }

                return _explorers;
            }
        }

        /// <summary>
        /// Returns the Microsoft Outlook version number.
        /// </summary>
        public Version Version
        {
            get
            {
                return new Version(InternalObject.GetType().InvokeMember("Version", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString());
            }
        }

        public object Run(string proc, params object[] args)
        {
            List<object> inArgs = new List<object>(args);
            inArgs.Insert(0, proc);

            return InternalObject.GetType().InvokeMember("Run", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, inArgs.ToArray());
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                if (_explorers != null)
                {
                    _explorers.Dispose();
                    _explorers = null;
                }
            }

            base.Dispose(true);
        }

        #endregion
        
    }
}
