using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Excel.Enums;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Represents the page setup description. The PageSetup object contains all page setup attributes (left margin, bottom margin, paper size, and so on) as properties.
    /// </summary>
    [WrapsCOM("Excel.PageSetup", "000208B4-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class PageSetup : ComWrapper
    {
        private Application _application;

        internal PageSetup(object interiorObj)
            : base(interiorObj)
        {
        }


        /// <summary>
        /// Gets a reference to the currently worksheet to which the range belongs. The returned worksheet is internally cached and does not need to be manually disposed.
        /// </summary>
        public Application Application
        {
            get
            {
                if (_application == null)
                    _application = new Application(InternalObject.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));

                return _application;
            }
        }


        /// <summary>
        /// Gets or sets a PageOrientation value that represents the protrait or landscape printing mode. (Requires Excel 2007)
        /// </summary>
        public PageOrientation Orientation
        {
            //TODO figure out if this works on previous versions of Excel. If not, either throw a friendly exception, or implement the equivilent 2003 code
            get
            {
                return (PageOrientation)InternalObject.GetType().InvokeMember("Orientation", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("Orientation", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the size of the top margin, in points. Read/write Double.
        /// </summary>
        public double TopMargin
        {
            get
            {
                return (double)InternalObject.GetType().InvokeMember("TopMargin", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("TopMargin", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the size of the bottom margin, in points. Read/write Double.
        /// </summary>
        public double BottomMargin
        {
            get
            {
                return (double)InternalObject.GetType().InvokeMember("BottomMargin", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("BottomMargin", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the size of the left margin, in points. Read/write Double.
        /// </summary>
        public double LeftMargin
        {
            get
            {
                return (double)InternalObject.GetType().InvokeMember("LeftMargin", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("LeftMargin", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the size of the right margin, in points. Read/write Double.
        /// </summary>
        public double RightMargin
        {
            get
            {
                return (double)InternalObject.GetType().InvokeMember("RightMargin", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("RightMargin", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the size of the footer margin, in points. Read/write Double.
        /// </summary>
        public double FooterMargin
        {
            get
            {
                return (double)InternalObject.GetType().InvokeMember("FooterMargin", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("FooterMargin", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the size of the header margin, in points. Read/write Double.
        /// </summary>
        public double HeaderMargin
        {
            get
            {
                return (double)InternalObject.GetType().InvokeMember("HeaderMargin", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("HeaderMargin", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the text for the left footer
        /// </summary>
        public string LeftFooter
        {
            get 
            {
                return InternalObject.GetType().InvokeMember("LeftFooter", System.Reflection.BindingFlags.GetProperty, null, InternalObject,  null).ToString();
            }

            set
            {
                InternalObject.GetType().InvokeMember("LeftFooter", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the text for the center footer
        /// </summary>
        public string CenterFooter
        {
            get
            {
                return InternalObject.GetType().InvokeMember("CenterFooter", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString();
            }

            set
            {
                InternalObject.GetType().InvokeMember("CenterFooter", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        /// <summary>
        /// Gets or sets the text for the right footer
        /// </summary>
        public string RightFooter
        {
            get
            {
                return InternalObject.GetType().InvokeMember("RightFooter", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null).ToString();
            }

            set
            {
                InternalObject.GetType().InvokeMember("RightFooter", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

    }
}
