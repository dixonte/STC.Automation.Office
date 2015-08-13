using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace MessageFilter
{
    public delegate void CalleeBusyHandler(object sender, CalleeBusyEventArgs args);
    public class CalleeBusyEventArgs : EventArgs
    {
        public CalleeBusyEventArgs() : base()
        {
            RetryDelay = 0;
            Cancel = false;
        }

        public int CalleeProcId { get; set; }
        public IntPtr TaskHandle { get; set; }
        public int TickCount { get; set; }

        // To be set by handler
        public int RetryDelay { get; set; }
        public bool Cancel { get; set; }
    }

    public class SimpleMessageFilter : IOleMessageFilter, IDisposable
    {
        // Start the filter.
        public SimpleMessageFilter()
        {
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(this, out oldFilter);
        }

        //
        // IOleMessageFilter functions.
        // Handle incoming thread requests.
        int IOleMessageFilter.HandleInComingCall(int dwCallType, System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr lpInterfaceInfo)
        {
            //Return the flag SERVERCALL_ISHANDLED.
            return 0;
        }

        public event CalleeBusyHandler CalleeBusy;

        // Thread call was rejected, so try again.
        int IOleMessageFilter.RetryRejectedCall(System.IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == 2)
            // flag = SERVERCALL_RETRYLATER.
            {
                if (CalleeBusy != null)
                {
                    IntPtr hThread = OpenThread((IntPtr)0x0040, false, hTaskCallee);
                    IntPtr procId = GetProcessIdOfThread(hThread);
                    CloseHandle(hThread);

                    var eventArgs = new CalleeBusyEventArgs() { CalleeProcId = (int)procId, TaskHandle = hTaskCallee, TickCount = dwTickCount };

                    CalleeBusy(this, eventArgs);

                    // Cancel retry
                    if (eventArgs.Cancel)
                        return -1;

                    return Math.Max(eventArgs.RetryDelay, 0); // Ensure values under 0 are not returned to COM, as we don't know what that will do except in the case of -1.
                }
                else
                {
                    // Retry the thread call immediately if return >=0 & <100.
                    return 0;
                }
            }
            // Too busy; cancel call.
            return -1;
        }

        int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            //Return the flag PENDINGMSG_WAITDEFPROCESS.
            return 2;
        }

        // Implement the IOleMessageFilter interface.
        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out  IOleMessageFilter oldFilter);
        [DllImport("kernel32")]
        private static extern IntPtr GetProcessIdOfThread(IntPtr threadId);
        [DllImport("kernel32")]
        private static extern IntPtr OpenThread(IntPtr dwDesiredAccess, bool bInheritHandle, IntPtr dwThreadId);
        [DllImport("kernel32")]
        private static extern bool CloseHandle(IntPtr hObject);
        [DllImport("user32.dll")]
        static extern IntPtr SetActiveWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        #region IDisposable Members

        protected void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free managed
                // Clear events, since GC apparently doesn't handle them well
                CalleeBusy = null;
            }

            // Free unmanaged
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(null, out oldFilter);
        }

        public void Dispose()
        {
            Dispose(true);
        }

        ~SimpleMessageFilter()
        {
            Dispose(false);
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
