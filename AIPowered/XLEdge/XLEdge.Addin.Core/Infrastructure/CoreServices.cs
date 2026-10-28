using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using XLEdge.Contracts;

namespace XLEdge.Addin.Core.Infrastructure
{
    /// <summary>
    /// What the core holds for its lifetime: the shell's context, and the core's own Excel RCW. The RCW is built from the
    /// shell's IDispatch pointer (same process), so every COM call from ported code goes straight to Excel rather than through
    /// a remoting proxy of the shell's RCW.
    /// <para>
    /// It is created in <see cref="Initialize"/>, which runs on Excel's main thread, so it is bound to that thread as today's
    /// HostApplication was. It is released only at <see cref="Shutdown"/>: ported code holds "var app = XLApp.App" locals
    /// (and every app.Workbooks... result shares the same RCW), so a dropped RCW (<see cref="ResetExcel"/>) is kept until then
    /// rather than destroyed under them. A reference kept past Shutdown would keep EXCEL.EXE alive.
    /// </para>
    /// </summary>
    internal static class CoreServices
    {
        // RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, CO_E_OBJNOTCONNECTED: the object is gone. Busy answers (RPC_E_CALL_REJECTED,
        // RPC_E_SERVERCALL_RETRYLATER - Excel editing a cell or showing a dialog) are not.
        private static readonly HashSet<int> DisconnectedResults = new HashSet<int>
        {
            unchecked((int)0x80010108), unchecked((int)0x800706BA), unchecked((int)0x800401FD),
        };

        private static readonly object Gate = new object();
        private static readonly List<object> Dropped = new List<object>();
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
                    if (_excel == null && _context != null)
                    {
                        _excel = CreateExcel(_context);
                        // Only after a ResetExcel: the new RCW is bound to this thread, which may not be Excel's.
                        if (_excel != null) Log(l => l.LogDebug(
                            $"CoreServices: Excel re-acquired on thread {Thread.CurrentThread.ManagedThreadId} ({Thread.CurrentThread.GetApartmentState()})."));
                    }
                    return _excel;
                }
            }
        }

        /// <summary>The Excel RCW if one exists; never creates one (today's ExcelAppUnsafe semantics).</summary>
        public static object ExcelAppIfCreated
        {
            get { lock (Gate) return _excel; }
        }

        /// <summary>
        /// Stops using the current Excel RCW (it was disconnected); the next <see cref="ExcelApp"/> asks the shell again. The
        /// dropped RCW is not released - code may still hold it - until <see cref="Shutdown"/>.
        /// </summary>
        public static void ResetExcel()
        {
            lock (Gate)
            {
                if (_excel != null) Dropped.Add(_excel);
                _excel = null;
            }
        }

        /// <summary>True when <paramref name="ex"/> means the COM object is gone, not merely busy.</summary>
        public static bool IsDisconnected(Exception ex) =>
            ex is InvalidComObjectException || (ex is COMException com && DisconnectedResults.Contains(com.HResult));

        /// <summary>Called by AddinEntry.Initialize on Excel's main thread: takes the context and Excel at once.</summary>
        public static void Initialize(IXLEdgeContext context)
        {
            lock (Gate)
            {
                ReleaseAll();
                _context = context;
                _excel = context == null ? null : CreateExcel(context);
            }
        }

        public static void Shutdown()
        {
            lock (Gate)
            {
                try { ReleaseAll(); }
                finally { _context = null; }
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
                Log(l => l.LogWarn("CoreServices: could not get Excel from the shell: " + ex.Message), context);
                return null;
            }
            finally
            {
                // The RCW holds its own reference; this one was AddRef'ed for us by the shell.
                if (dispatch != IntPtr.Zero) Marshal.Release(dispatch);
            }
        }

        private static void ReleaseAll()
        {
            var all = new List<object>(Dropped);
            if (_excel != null) all.Add(_excel);
            Dropped.Clear();
            _excel = null;
            foreach (object excel in all)
            {
                try
                {
                    if (Marshal.IsComObject(excel)) Marshal.FinalReleaseComObject(excel);
                }
                catch (Exception)
                {
                    // Already released by ported code (today's shutdown released it itself) or gone with Excel.
                }
            }
        }

        private static void Log(Action<ILogger> write, IXLEdgeContext context = null)
        {
            try
            {
                ILogger logger = (context ?? _context)?.Logger;
                if (logger != null) write(logger);
            }
            catch (Exception)
            {
                // Logging must not fail the call.
            }
        }
    }
}
