using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace StreamDeckVS
{
    public static class WindowsAPI
    {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern IntPtr GetParent(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hwnd, int flags);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern uint GetWindowModuleFileNameA(IntPtr hwnd, StringBuilder lpszFileName, uint cchFileNameMax);

        //static extern IntPtr GetAncestor(IntPtr hwnd, GetAncestorFlags flags);
        //GetAncestor(hWnd, GA_PARENT);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("ole32.dll")]
        public static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        [DllImport("ole32.dll")]
        public static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);
    }
}
