using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;

namespace STC.Automation.Office.Common
{
    /// <summary>
    /// Abstract class for wrapping a COM object event sinks. Calls ConnetionPoint.Unadvise and Marshal.ReleaseComObject when disposed or destroyed.
    /// </summary>
    public abstract class ComEventSinkWrapper<I> : IDisposable
        where I: class
    {
        /// <summary>
        /// The object which owns this event sink. Used as the 'sender' of events.
        /// </summary>
        protected WeakReference Parent { get; private set; }

        private IConnectionPoint _connectionPoint;
        private int _sinkCookie;

        internal ComEventSinkWrapper(ComWrapper parent)
        {
            _sinkCookie = -1;

            Guid guid = typeof(I).GUID;
            try
            {
                ((IConnectionPointContainer)parent.InternalObject).FindConnectionPoint(ref guid, out _connectionPoint);
            }
            catch
            {
                throw new COMException(string.Format("Could not wrap event sink for {0}. Could not find connection point.", typeof(I).Name));
            }

            if (_connectionPoint == null)
            {
                throw new COMException(string.Format("Could not wrap event sink for {0}. Could not find connection point.", typeof(I).Name));
            }

            Parent = new WeakReference(parent);
            _connectionPoint.Advise(this, out _sinkCookie);
        }

        /// <summary>
        /// Destroys this event sink wrapper.
        /// </summary>
        ~ComEventSinkWrapper()
        {
            Dispose(false);
        }

        internal virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
            }

            if (_connectionPoint != null)
            {
                if (_sinkCookie >= 0)
                    _connectionPoint.Unadvise(_sinkCookie);

                Marshal.ReleaseComObject(_connectionPoint);
                _connectionPoint = null;
            }
        }

        #region IDisposable Members

        /// <summary>
        /// Cleans up this event sink wrapper.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
