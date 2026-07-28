using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Helper for cancellation token operations with detailed logging
    /// </summary>
    public static class CancellationTokenHelper
    {
        /// <summary>
        /// Throws if cancellation requested with detailed logging
        /// </summary>
        public static void ThrowIfCancelled(
            CancellationToken token,
            [CallerMemberName] string operationName = "")
        {
            if (token.IsCancellationRequested)
            {
                LogUtility.LogWarn($"Operation cancelled: {operationName}");
                token.ThrowIfCancellationRequested();
            }
        }

        /// <summary>
        /// Checks if cancellation is requested and logs
        /// </summary>
        public static bool IsCancellationRequested(
            CancellationToken token,
            [CallerMemberName] string operationName = "")
        {
            if (token.IsCancellationRequested)
            {
                LogUtility.LogWarn($"Cancellation detected in: {operationName}");
                return true;
            }

            return false;
        }

        /// <summary>
        /// Waits asynchronously with cancellation support and logging
        /// </summary>
        public static async Task DelayWithLogging(
            int milliseconds,
            CancellationToken token,
            string reason = "")
        {
            try
            {
                LogUtility.LogDebug($"Waiting {milliseconds}ms{(string.IsNullOrEmpty(reason) ? "" : $" - {reason}")}");

                await Task.Delay(milliseconds, token);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn($"Wait cancelled after delay{(string.IsNullOrEmpty(reason) ? "" : $": {reason}")}");
                throw;
            }
        }
    }
}
