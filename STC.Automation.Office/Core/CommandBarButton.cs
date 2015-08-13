using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using STC.Automation.Office.Common;
using STC.Automation.Office.Attributes;
using STC.Automation.Office.Word.Enums;
using System.Drawing;

namespace STC.Automation.Office.Core
{
    /// <summary>
    /// Wraps an Core.CommandBarButton COM object
    /// </summary>
    [WrapsCOM("Core.CommandBarButton", CommandBarButton.UUID)]
    public class CommandBarButton : CommandBarControl
    {
        public const string UUID = "000C030E-0000-0000-C000-000000000046";

        /// <summary>
        /// Wraps the given COM object as a CommandBarButton.
        /// </summary>
        /// <param name="commandBarButtonObj"></param>
        public CommandBarButton(object commandBarButtonObj)
            : base(commandBarButtonObj)
        {
        }

        public bool BuiltInFace
        {
            get
            {
                return (bool)InternalObject.GetType().InvokeMember("BuiltInFace", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("BuiltInFace", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        public Image Picture
        {
            get
            {
                // TODO: Insert the following into host application via VBE:
                //Public Function GetPictureBytes(ByVal commandbutton As Object) As Byte()
                //    Dim helper As Object
                //    Set helper = CreateObject("IPictureDispHelper.PictDispHelper")

                //    GetPictureBytes = helper.GetBytes(commandbutton.Picture)
                //End Function

                //Public Function GetMaskBytes(ByVal parent As Object) As Byte()
                //    Dim helper As Object
                //    Set helper = CreateObject("IPictureDispHelper.PictDispHelper")

                //    GetMaskBytes = helper.GetBytes(commandbutton.Mask)
                //End Function

                //object res = Application.Run("GetPictureBytes", InternalObject);

                //if (res is byte[])
                //{
                //    var data = (byte[])res;
                //    if (data.Length > 0)
                //    {
                //        using (var mem = new System.IO.MemoryStream(data))
                //        {
                //            return Image.FromStream(mem);
                //        }
                //    }
                //    else
                //    {
                //        return null;
                //    }
                //}

                return null;
            }

            set
            {
            }
        }

        public string OnAction
        {
            get
            {
                return (string)InternalObject.GetType().InvokeMember("OnAction", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("OnAction", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        public string Parameter
        {
            get
            {
                return (string)InternalObject.GetType().InvokeMember("Parameter", System.Reflection.BindingFlags.GetProperty, null, InternalObject, null);
            }

            set
            {
                InternalObject.GetType().InvokeMember("Parameter", System.Reflection.BindingFlags.SetProperty, null, InternalObject, new object[] { value });
            }
        }

        #region ComWrapper Members

        internal override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
            }

            base.Dispose(true);
        }

        #endregion
    }
}
