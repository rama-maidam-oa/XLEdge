using System;
using System.Runtime.InteropServices;
using XLEdge.Contracts;

namespace XLEdge.Addin.Core.Infrastructure
{
    /// <summary>
    /// What the core holds for its lifetime: the shell's context, and the core's own Excel RCW. The RCW is built from the
    /// shell's IDispatch pointer (same process), so every COM call from ported code goes straight to Excel rather than through
    /// a remoting proxy of the shell's RCW. Released at Shutdown: a reference kept here would keep EXCEL.EXE alive.
    /// </summary>
    internal static class CoreServices
    {
        private static readonly object Gate = new object();
        private static IXLEdgeContext _context;
        private static object _excel;

        /// <summary>The shell's context; null before Initialize and after Shutdown.</summary>
        public static IXLEdgeContext Context
        {
            get { lock (Gate) return _context; }
        }

        /// <summary>The Excel Application (a COM object of this AppDomain), or null when there is none or no context.</summary>
        public static object ExcelApp
        {
            get
            {
                lock (Gate)
                {
                    if (_excel == null && _context != null) _excel = CreateExcel(_context);
                    return _excel;
                }
            }
        }

        public static void Initialize(IXLEdgeContext context)
        {
            lock (Gate)
            {
                ReleaseExcel();
                _context = context;
            }
        }

        public static void Shutdown()
        {
            lock (Gate)
            {
                ReleaseExcel();
                _context = null;
            }
        }

        private static object CreateExcel(IXLEdgeContext context)
        {
            IntPtr dispatch = IntPtr.Zero;
            try
            {
                dispatch = context.GetExcelDispatch();
                return dispatch == IntPtr.Zero ? null : Marshal.GetObjectForIUnknown(dispatch);
            }
            catch (Exception ex)
            {
                try { context.Logger?.LogWarn("CoreServices: could not get Excel from the shell: " + ex.Message); }
                catch (Exception) { /* logging must not fail the call */ }
                return null;
            }
            finally
            {
                // The RCW holds its own reference; this one was AddRef'ed for us by the shell.
                if (dispatch != IntPtr.Zero) Marshal.Release(dispatch);
            }
        }

        private static void ReleaseExcel()
        {
            object excel = _excel;
            _excel = null;
            if (excel != null && Marshal.IsComObject(excel)) Marshal.FinalReleaseComObject(excel);
        }
    }
}
