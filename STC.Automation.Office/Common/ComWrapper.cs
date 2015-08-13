using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using STC.Automation.Office.Attributes;
using System.Runtime.InteropServices.ComTypes;
using System.Reflection;

namespace STC.Automation.Office.Common
{
    /// <summary>
    /// Abstract class for wrapping a COM object. Calls Marshal.ReleaseComObject when disposed or destroyed.
    /// </summary>
    public abstract class ComWrapper : IDisposable
    {
        private object _wrappedObject;

        public bool IsDisposed
        {
            get;
            private set;
        }

        internal ComWrapper()
        {
            Type progType;
            if (!String.IsNullOrEmpty(WrapsProgId) && (((progType = Type.GetTypeFromProgID(WrapsProgId)) != null)))
            {
                Wrap(Activator.CreateInstance(progType));
            }

            if (_wrappedObject == null)
            {
                throw new COMException("Could not create instance of " + WrapsProgId);
            }
        }

        internal ComWrapper(object objToWrap)
        {
            Wrap(objToWrap);
        }


        private void Wrap(object objToWrap)
        {
            if ((objToWrap != null) && Marshal.IsComObject(objToWrap))
            {
                if (MustSupport != null)
                {
                    foreach (Guid req in MustSupport)
                    {
                        if (!SupportsInterface(objToWrap, req))
                        {
                            throw new COMException(string.Format("Problem wrapping {0} object; does not support interface {{{1}}}.", WrapsProgId, req.ToString()));
                        }
                    }
                }

                _wrappedObject = objToWrap;

                // If a session is active, add to session
                Session.AddWrapper(this);
            }
            else
            {
                throw new COMException("Cannot wrap null or non-COM object.");
            }
        }

        internal string WrapsProgId
        {
            get
            {
                string progId = "UNKNOWN";

                object[] attribs = this.GetType().GetCustomAttributes(typeof(WrapsCOMAttribute), true);
                if (attribs.Length > 0)
                    progId = ((WrapsCOMAttribute)attribs[0]).ProgID;

                return progId;
            }
        }

        internal Guid[] MustSupport
        {
            get
            {
                return GetMustSupport(this.GetType());
            }
        }

        internal static Guid[] GetMustSupport(Type type)
        {
            Guid[] mustSupport = null;

            object[] attribs = type.GetCustomAttributes(typeof(WrapsCOMAttribute), true);
            if (attribs.Length > 0)
                mustSupport = ((WrapsCOMAttribute)attribs[0]).MustSupport;

            return mustSupport;
        }

        internal static bool SupportsInterface(object obj, Guid[] ifaces)
        {
            foreach (var i in ifaces)
                if (SupportsInterface(obj, i))
                    return true;

            return false;
        }

        internal static bool SupportsInterface(object obj, Guid iface)
        {
            IntPtr iunknown, ifaceRef;

            iunknown = Marshal.GetIUnknownForObject(obj);
            if (iunknown != IntPtr.Zero)
            {
                try
                {
                    Marshal.QueryInterface(iunknown, ref iface, out ifaceRef);

                    if (ifaceRef != IntPtr.Zero)
                    {
                        Marshal.Release(ifaceRef);

                        return true;
                    }
                }
                finally
                {
                    Marshal.Release(iunknown);
                }
            }

            return false;
        }

        [DllImport("User32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("User32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildCallback lpEnumFunc, ref IntPtr lParam);

        [DllImport("Oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, byte[] riid, [MarshalAs(UnmanagedType.IDispatch)] ref object ptr);

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        private delegate bool EnumChildCallback(IntPtr hwnd, ref IntPtr lParam);
        internal static T FromProcess<T>(Process processToAttach, string className)
            where T: ComWrapper
        {
            // First, get the main window handle.
            IntPtr hwnd = processToAttach.MainWindowHandle;

            // We need to enumerate the child windows to find one that
            // supports accessibility. To do this, instantiate the
            // delegate and wrap the callback method in it, then call
            // EnumChildWindows, passing the delegate as the 2nd arg.
            if (hwnd != IntPtr.Zero)
            {
                IntPtr hwndChild = IntPtr.Zero;
                var cb = new EnumChildCallback(
                    (IntPtr _hwnd, ref IntPtr _lParam)
                =>
                    {
                        StringBuilder buf = new StringBuilder(128);
                        GetClassName(_hwnd, buf, 128);
                        if (buf.ToString() == className)
                        {
                            _lParam = _hwnd;
                            return false;
                        }
                        return true;
                    });
                EnumChildWindows(hwnd, cb, ref hwndChild);

                // If we found an accessible child window, call
                // AccessibleObjectFromWindow, passing the constant
                // OBJID_NATIVEOM (defined in winuser.h) and
                // IID_IDispatch - we want an IDispatch pointer
                // into the native object model.
                if (hwndChild != IntPtr.Zero)
                {
                    const uint OBJID_NATIVEOM = 0xFFFFFFF0;
                    Guid IID_IDispatch = new Guid(
                         "{00020400-0000-0000-C000-000000000046}");
                    object ptr = null;

                    int hr = AccessibleObjectFromWindow(hwndChild, OBJID_NATIVEOM, IID_IDispatch.ToByteArray(), ref ptr);

                    if (hr >= 0)
                    {
                        // If we successfully got a native OM
                        // IDispatch pointer, we can QI this for
                        // an Office Application (using the implicit
                        // cast operator supplied in the PIA).
                        return Activator.CreateInstance(typeof(T), new object[] { ptr }) as T;
                    }
                }
            }

            return null;
        }

        internal static IList<T> FromROT<T>()
            where T: ComWrapper
        {
            var result = new List<T>();

            IntPtr numFetched = IntPtr.Zero;
            IRunningObjectTable runningObjectTable;
            IEnumMoniker monikerEnumerator;
            IMoniker[] monikers = new IMoniker[1];

            GetRunningObjectTable(0, out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                IBindCtx ctx;
                CreateBindCtx(0, out ctx);

                //string runningObjectName;
                //monikers[0].GetDisplayName(ctx, null, out runningObjectName);

                object runningObjectVal;
                runningObjectTable.GetObject(monikers[0], out runningObjectVal);
                
                if (runningObjectVal == null)
                    continue;

                bool supportsAll = true;
                foreach (var guid in GetMustSupport(typeof(T)))
                {
                    if (!SupportsInterface(runningObjectVal, guid))
                    {
                        supportsAll = false;
                        break;
                    }
                }

                if (supportsAll)
                {
                    var constructor = typeof(T).GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public, null, new Type[] { typeof(Object) }, null);
                    
                    result.Add((T)constructor.Invoke(new object[] { runningObjectVal }));
                }
                else
                {
                    Marshal.ReleaseComObject(runningObjectVal);
                }
            }

            return result;
        }

        /// <summary>
        /// Destroys this COM wrapper.
        /// </summary>
        ~ComWrapper()
        {
            Dispose(false);
        }

        /// <summary>
        /// The unwrapped form of the COM object this class wraps.
        /// </summary>
        public object InternalObject
        {
            get
            {
                if (!IsDisposed)
                {
                    return _wrappedObject;
                }
                else
                {
                    throw new ObjectDisposedException(this.GetType().Name);
                }
            }
        }

        #region IDisposable Members

        internal virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                IsDisposed = true;
            }

            if (_wrappedObject != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_wrappedObject);
                }
                catch (COMException) { }  //TODO this is here because Adam has an urgent project that needs to be done and the above line of code sometimes crashes so that his project cannot proceed. There may be a better way to fix this, and that should be investigated at some point.
                _wrappedObject = null;
            }
        }

        /// <summary>
        /// Cleans up this COM wrapper.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
