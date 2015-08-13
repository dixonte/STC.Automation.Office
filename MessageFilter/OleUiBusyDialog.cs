using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MessageFilter
{
    public static class OleUiBusyDialog
    {
        private enum BZFlags
        {
            DisableCancelButton = 0x00000001,
            DisableSwitchToButton = 0x00000002,
            DisableRetryButton = 0x00000004,
            NotRespondingDialog = 0x00000008
        }

        [StructLayout(LayoutKind.Sequential, Pack=1)]
        private struct tagOLEUIBUSY
        {
            public UInt32 cbStruct;
            public Int32 dwFlags;
            public IntPtr hWndOwner;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpszCaption;
            public IntPtr lpfnHook;
            public IntPtr lCustData;
            public IntPtr hInstance;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpszTemplate;
            public IntPtr hResource;
            public IntPtr hTask;
            public IntPtr lphWndDialog;
        }

        public enum OLEUIFlags
        {
            OLEUI_FALSE = 0,
            OLEUI_SUCCESS = 1,      // No error, same as OLEUI_OK
            OLEUI_OK = 1,      // OK button pressed
            OLEUI_CANCEL = 2,      // Cancel button pressed

            OLEUI_ERR_STANDARDMIN = 100,
            OLEUI_ERR_OLEMEMALLOC = 100,
            OLEUI_ERR_STRUCTURENULL = 101,    // Standard field validation
            OLEUI_ERR_STRUCTUREINVALID = 102,
            OLEUI_ERR_CBSTRUCTINCORRECT = 103,
            OLEUI_ERR_HWNDOWNERINVALID = 104,
            OLEUI_ERR_LPSZCAPTIONINVALID = 105,
            OLEUI_ERR_LPFNHOOKINVALID = 106,
            OLEUI_ERR_HINSTANCEINVALID = 107,
            OLEUI_ERR_LPSZTEMPLATEINVALID = 108,
            OLEUI_ERR_HRESOURCEINVALID = 109,

            OLEUI_ERR_FINDTEMPLATEFAILURE = 110,    // Initialization errors
            OLEUI_ERR_LOADTEMPLATEFAILURE = 111,
            OLEUI_ERR_DIALOGFAILURE = 112,
            OLEUI_ERR_LOCALMEMALLOC = 113,
            OLEUI_ERR_GLOBALMEMALLOC = 114,
            OLEUI_ERR_LOADSTRING = 115,

            OLEUI_ERR_STANDARDMAX = 116,

            OLEUI_BZERR_HTASKINVALID = 116,
            OLEUI_BZ_SWITCHTOSELECTED = 117,
            OLEUI_BZ_RETRYSELECTED = 118,
            OLEUI_BZ_CALLUNBLOCKED = 119
        }

        [DllImport("OleDlg.dll")]
        private static extern OLEUIFlags OleUIBusy([MarshalAs(UnmanagedType.Struct), In] ref tagOLEUIBUSY lpBZ);

        public static OLEUIFlags Show(IWin32Window owner, IntPtr hTask, string caption)
        {
            var lpBZ = new tagOLEUIBUSY();

            lpBZ.lpszCaption = caption;
            lpBZ.hWndOwner = owner.Handle;
            lpBZ.hTask = hTask;
            lpBZ.cbStruct = (UInt32)Marshal.SizeOf(lpBZ);

            return OleUIBusy(ref lpBZ);
        }
    }
}
