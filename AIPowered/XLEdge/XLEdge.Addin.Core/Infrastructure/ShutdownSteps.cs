using System;
using XLEdge.Utilities;

namespace XLEdge.Addin.Core.Infrastructure
{
    /// <summary>
    /// Runs shutdown steps in order, each in its own try: a failing step is logged and the next one still runs, so one
    /// failure (a pane's WebView2 that will not dispose) never leaves the later steps' windows, dispatcher or COM references
    /// alive into the AppDomain unload.
    /// </summary>
    internal static class ShutdownSteps
    {
        public static void RunAll(string owner, params (string Name, Action Step)[] steps)
        {
            foreach ((string name, Action step) in steps)
            {
                try { step(); }
                catch (Exception ex) { LogUtility.LogException(ex, owner + " (" + name + ")"); }
            }
        }
    }
}
