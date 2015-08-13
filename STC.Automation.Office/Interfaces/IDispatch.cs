using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace STC.Automation.Office.Interfaces
{
    /// <summary>
    /// Partial import of the COM interface IDispatch.
    /// </summary>
    [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDispatch
    {
        /// <summary>
        /// Reserved.
        /// </summary>
        void Reserved();
        
        /// <summary>
        /// Gets an ITypeInfo interface for the object.
        /// </summary>
        /// <param name="nInfo">The type information to return. Pass 0 to retrieve type information for the IDispatch implementation.</param>
        /// <param name="lcid">The locale identifier for the type information. An object may be able to return different type information for different languages. This is important for classes that support localized member names. For classes that do not support localized member names, this parameter can be ignored.</param>
        /// <param name="typeInfo">Receives the ITypeInfo interface.</param>
        /// <returns>HRESULT</returns>
        [PreserveSig]
        int GetTypeInfo(uint nInfo, int lcid, out ITypeInfo typeInfo);
    }

    /// <summary>
    /// Helper class for IDispatch
    /// </summary>
    public static class DispatchHelper
    {
        /// <summary>
        /// Gets the TypeName for objects implementing IDispatch.
        /// </summary>
        /// <param name="obj">The object for which to get the TypeName</param>
        /// <returns>String</returns>
        public static string GetIDispatchTypeName(IDispatch obj)
        {
            ITypeInfo t;

            obj.GetTypeInfo(0, 0, out t);

            return Marshal.GetTypeInfoName(t);
        }
    }
}
