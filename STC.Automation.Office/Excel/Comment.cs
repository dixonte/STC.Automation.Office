using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Excel.Enums;
using System.Drawing;
using System.Reflection;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Represents a cell comment.
    /// </summary>
    [WrapsCOM("Excel.Comment", "00024427-0000-0000-C000-000000000046")]
    [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class Comment : ComWrapper
    {
         private Application _application;

         internal Comment(object interiorObj)
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
         /// Returns or sets the author of the comment.
         /// </summary>
         [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
         public string Author
         {
             get
             {
                 return (string)InternalObject.GetType().InvokeMember("Author", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
             }

             set
             {
                 InternalObject.GetType().InvokeMember("Author", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
             }
         }

         /// <summary>
         /// Returns or sets a Boolean value that determines whether the object is visible.
         /// </summary>
         [System.Obsolete("This property has not been fully tested yet and is not guaranteed to work")]
         public bool Visible
         {
             get
             {
                 return (bool)InternalObject.GetType().InvokeMember("Visible", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
             }

             set
             {
                 InternalObject.GetType().InvokeMember("Author", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
             }
         }

         [System.Obsolete("This method has not been fully tested yet and is not guaranteed to work")]
         public void Text(string text)
         {
             Text(text, null, null);
         }

         /// <summary>
         /// Sets comment text.
         /// </summary>
         /// <param name="text">The text to be added.</param>
         /// <param name="start">The character number where the added text will be placed. If this argument is omitted, any existing text in the comment is deleted.</param>
         /// <param name="overwrite">True to overwrite the existing text. The default value is False (text is inserted).</param>
         [System.Obsolete("This method has not been fully tested yet and is not guaranteed to work")]
         public void Text(string text, int? start, bool? overwrite)
         {
             InternalObject.GetType().InvokeMember("Text", System.Reflection.BindingFlags.InvokeMethod, null, InternalObject, new object[] { (text == null ? Missing.Value : (object)text), (start.HasValue ? (object)start : Missing.Value), (overwrite.HasValue ? (object)overwrite : Missing.Value) });
         }
    }
}
