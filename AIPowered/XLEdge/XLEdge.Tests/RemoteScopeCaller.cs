using System;
using System.Threading.Tasks;
using XLEdge.Contracts;

namespace XLEdge.Tests
{
    /// <summary>
    /// Runs inside a second AppDomain and logs through the Logger proxy, the way the XLEdge core does.
    /// </summary>
    public sealed class RemoteScopeCaller : MarshalByRefObject
    {
        public void Run(ILogger logger, string scopeName)
        {
            using (logger.BeginLogScope(scopeName))
            {
                logger.LogDebug("inside");
            }
        }

        public void RunWithAwait(ILogger logger, string scopeName) => RunAsync(logger, scopeName).GetAwaiter().GetResult();

        private static async Task RunAsync(ILogger logger, string scopeName)
        {
            using (logger.BeginLogScope(scopeName))
            {
                await Task.Delay(20).ConfigureAwait(false);
                logger.LogDebug("inside");
            }
        }
    }
}
