using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Excel.Enums;
using System.Drawing;

namespace STC.Automation.Office.Excel
{
    /// <summary>
    /// Represents the interior of an object.
    /// </summary>
     [WrapsCOM("Excel.Interior", "00020870-0000-0000-C000-000000000046")]
     [System.Obsolete("This class has not been fully tested yet and it not guaranteed to work")]
    public class Interior : ComWrapper
    {
         private Application _application;

         internal Interior(object interiorObj)
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
         /// Returns or sets a Pattern enum, that represents the interior pattern.
         /// </summary>
         public Pattern Pattern
         {
             get
             {
                 return (Pattern)InternalObject.GetType().InvokeMember("Pattern", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
             }
             set
             {
                 InternalObject.GetType().InvokeMember("Pattern", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
             }
         }

         /// <summary>
         /// Gets or Sets a value that represents the color of the interior.
         /// This can be an index into the current colour pallette, or as an XLColorIndex
         /// </summary>
         public int ColorIndex
         {
             get
             {
                 return (int)InternalObject.GetType().InvokeMember("ColorIndex", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
             }

             set
             {
                 InternalObject.GetType().InvokeMember("ColorIndex", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
             }
         }



         /// <summary>
         /// Gets or Sets a value that represents the color of the interior.
         /// </summary>
         public Color Color
         {
             get
             {
                 return ColorTranslator.FromOle(Convert.ToInt32(InternalObject.GetType().InvokeMember("Color", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null)));
             }

             set
             {
                 InternalObject.GetType().InvokeMember("Color", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { ColorTranslator.ToOle(value) });
             }
         }
    }
}
