using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Windows;

namespace XLEdge.Utilities
{
    public static class UiDispatcher
    {
        public static Dispatcher Current =>
            Application.Current?.Dispatcher
            ?? throw new InvalidOperationException("Application dispatcher is not available.");

        public static Task RunAsync(Action action)
        {
            if (action == null)
                throw new ArgumentNullException(nameof(action));

            if (Current.CheckAccess())
            {
                action();
                return Task.CompletedTask;
            }

            return Current.InvokeAsync(action).Task;
        }

        /// <summary>
        /// Marshals an async delegate onto the UI dispatcher and awaits its full completion,
        /// including any inner <c>await</c>s (unlike the <see cref="RunAsync(Action)"/> overload above).
        /// </summary>
        public static async Task RunAsync(Func<Task> asyncAction)
        {
            if (asyncAction == null)
                throw new ArgumentNullException(nameof(asyncAction));

            if (Current.CheckAccess())
            {
                await asyncAction();
                return;
            }

            await await Current.InvokeAsync(asyncAction).Task;
        }

        public static Task<T> RunAsync<T>(Func<T> func)
        {
            if (func == null)
                throw new ArgumentNullException(nameof(func));

            if (Current.CheckAccess())
                return Task.FromResult(func());

            return Current.InvokeAsync(func).Task;
        }

        public static void Run(Action action)
        {
            if (action == null)
                throw new ArgumentNullException(nameof(action));

            if (Current.CheckAccess())
                action();
            else
                Current.Invoke(action);
        }
    }
}
