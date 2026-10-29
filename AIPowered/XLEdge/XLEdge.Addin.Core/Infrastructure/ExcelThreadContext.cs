using System;
using System.Threading;
using System.Windows.Threading;

namespace XLEdge.Addin.Core.Infrastructure
{
    /// <summary>
    /// The synchronization context 11.1.2's handlers run under in the core. XLEdge 11.1.2 ran them on Excel's thread under the
    /// shell's WinForms context, so every continuation after an await came back to Excel's thread; a call arriving from the
    /// shell's AppDomain brings no context into this one, and without it those continuations would run on pool threads and call
    /// Excel from there. This context posts them to the core's dispatcher on Excel's thread (pumped by Excel's message loop), and
    /// counts the async-void handlers still running, so the core reports itself busy until each one has ended.
    /// </summary>
    internal sealed class ExcelThreadContext : SynchronizationContext
    {
        private static int _pending;
        private readonly Dispatcher _dispatcher;

        private ExcelThreadContext(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
        }

        /// <summary>Async-void handlers started under this context that have not ended yet.</summary>
        public static int Pending => Volatile.Read(ref _pending);

        public override void OperationStarted() => Interlocked.Increment(ref _pending);

        public override void OperationCompleted() => Interlocked.Decrement(ref _pending);

        public override void Post(SendOrPostCallback d, object state) => _dispatcher.BeginInvoke(DispatcherPriority.Normal, d, state);

        public override void Send(SendOrPostCallback d, object state)
        {
            if (_dispatcher.CheckAccess()) d(state);
            else _dispatcher.Invoke(DispatcherPriority.Send, d, state);
        }

        public override SynchronizationContext CreateCopy() => new ExcelThreadContext(_dispatcher);

        /// <summary>Runs <paramref name="action"/> on the calling (Excel's) thread under this context, then restores the previous one.</summary>
        public static void Run(Action action)
        {
            SynchronizationContext previous = SynchronizationContext.Current;
            try
            {
                if (!(previous is ExcelThreadContext))
                    SynchronizationContext.SetSynchronizationContext(new ExcelThreadContext(Dispatcher.CurrentDispatcher));
                action();
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previous);
            }
        }
    }
}
