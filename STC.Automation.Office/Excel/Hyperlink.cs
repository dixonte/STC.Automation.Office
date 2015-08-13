using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Excel.Enums;
using System.Drawing;
using System.Reflection;
using STC.Automation.Office.Core.Enums;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Represents a hyperlink.
    /// </summary>
    [WrapsCOM("Excel.Hyperlink", "00024431-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class Hyperlink : ComWrapper
    {
         private Application _application;

         internal Hyperlink(object interiorObj)
             : base(interiorObj)
        {
        }


         /// <summary>
         /// Gets an Application object that represents the creator of the specified object. The returned Application is internally cached and does not need to be manually disposed.
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
         /// Returns or sets a String value that represents the address of the target document.
         /// </summary>
         [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
         public string Address
         {
             get
             {
                 return (string)InternalObject.GetType().InvokeMember("Address", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
             }

             set
             {
                 InternalObject.GetType().InvokeMember("Address", System.Reflection.BindingFlags.SetProperty, null, InternalObject, ComArguments.Prepare(value));
             }
         }

         /// <summary>
         /// Returns or sets the text string of the specified hyperlink’s e-mail subject line. The subject line is appended to the hyperlink’s address. Read/write String.
         /// </summary>
         [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
         public string EmailSubject
         {
             get
             {
                 return (string)InternalObject.GetType().InvokeMember("EmailSubject", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
             }

             set
             {
                 InternalObject.GetType().InvokeMember("EmailSubject", System.Reflection.BindingFlags.SetProperty, null, InternalObject, ComArguments.Prepare(value));
             }
         }

         /// <summary>
         /// Returns a String value that represents the name of the object.
         /// </summary>
         [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
         public string Name
         {
             get
             {
                 return (string)InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
             }

             set
             {
                 InternalObject.GetType().InvokeMember("Name", System.Reflection.BindingFlags.SetProperty, null, InternalObject, ComArguments.Prepare(value));
             }
         }

         /// <summary>
         /// Returns a Range object that represents the range the specified hyperlink is attached to.  Returned object must be manually disposed.
         /// </summary>
         [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
         public Range Range
         {
             get
             {
                 return new Range(InternalObject.GetType().InvokeMember("Range", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
             }

             set
             {
                 InternalObject.GetType().InvokeMember("Range", System.Reflection.BindingFlags.SetProperty, null, InternalObject, ComArguments.Prepare(value));
             }
         }

         /// <summary>
         /// Returns or sets the ScreenTip text for the specified hyperlink. Read/write String.
         /// </summary>
         [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
         public string ScreenTip
         {
             get
             {
                 return (string)InternalObject.GetType().InvokeMember("ScreenTip", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
             }

             set
             {
                 InternalObject.GetType().InvokeMember("ScreenTip", System.Reflection.BindingFlags.SetProperty, null, InternalObject, ComArguments.Prepare(value));
             }
         }

         /// <summary>
         /// Returns a Shape object that represents the shape attached to the specified hyperlink.  Returned object must be manually disposed.
         /// </summary>
         [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
         public Shape Shape
         {
             get
             {
                 return new Shape(InternalObject.GetType().InvokeMember("Shape", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null));
             }

             set
             {
                 InternalObject.GetType().InvokeMember("Shape", System.Reflection.BindingFlags.SetProperty, null, InternalObject, ComArguments.Prepare(value));
             }
         }

         /// <summary>
         /// Returns or sets the location within the document associated with the hyperlink. Read/write String.
         /// </summary>
         [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
         public string SubAddress
         {
             get
             {
                 return (string)InternalObject.GetType().InvokeMember("SubAddress", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
             }

             set
             {
                 InternalObject.GetType().InvokeMember("SubAddress", System.Reflection.BindingFlags.SetProperty, null, InternalObject, ComArguments.Prepare(value));
             }
         }

         /// <summary>
         /// Returns or sets the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink. Read/write String.
         /// </summary>
         [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
         public string TextToDisplay
         {
             get
             {
                 return (string)InternalObject.GetType().InvokeMember("TextToDisplay", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
             }

             set
             {
                 InternalObject.GetType().InvokeMember("TextToDisplay", System.Reflection.BindingFlags.SetProperty, null, InternalObject, ComArguments.Prepare(value));
             }
         }

         /// <summary>
         /// Returns a HyperlinkType value, that represents the location of the HTML frame.
         /// </summary>
         [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
         public HyperlinkType Type
         {
             get
             {
                 return (HyperlinkType)InternalObject.GetType().InvokeMember("Type", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
             }
         }

    }
}
