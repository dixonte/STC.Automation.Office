using System;
using System.Collections.Generic;
using System.Text;
using STC.Automation.Office.Common;

namespace STC.Automation.Office
{
    /// <summary>
    /// Create an instance of this class before doing any automation, and all Automation objects wrapped will be added to the session. Dispose this instance to dispose all wrapped COM objects created in that session.
    /// At this stage, you can only create one session. Sessions-within-sessions could be implemented later.
    /// </summary>
    public class Session: IDisposable
    {
        private static List<ComWrapper> s_wrappers;

        public Session()
        {
            if (s_wrappers != null)
                throw new InvalidOperationException("A session is already active!");

            s_wrappers = new List<ComWrapper>();
        }

        ~Session()
        {
            Dispose(false);
        }

        internal static void AddWrapper(ComWrapper wrapper)
        {
            if (s_wrappers != null)
                s_wrappers.Add(wrapper);
        }

        private void CloseSession()
        {
            if (s_wrappers != null)
            {
                foreach (var wrapper in s_wrappers)
                {
                    if (!wrapper.IsDisposed)
                        wrapper.Dispose();
                }

                s_wrappers.Clear();
                s_wrappers = null;
            }
        }

        #region IDisposable Members

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed, null large fields
                CloseSession();
            }

            // Free unmanaged
        }

        public void Dispose()
        {
            Dispose(true);
        }

        #endregion
    }
}
