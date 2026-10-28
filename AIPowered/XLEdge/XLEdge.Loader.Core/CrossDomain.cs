using System;
using System.Collections.Generic;
using System.Reflection;

namespace XLEdge.Loader.Core
{
    /// <summary>
    /// Keeps exceptions that cross from the core's AppDomain back into the shell's deserializable there: an exception whose type
    /// (or any inner exception's type) is not a .NET Framework type would need the core's assemblies in the shell's domain.
    /// </summary>
    internal static class CrossDomain
    {
        public static bool NeedsWrapping(Exception exception)
        {
            foreach (Exception e in Chain(exception))
                if (!e.GetType().Assembly.GlobalAssemblyCache) return true;
            return false;
        }

        /// <summary>A framework-only exception carrying the full text (types, messages, stack traces) of the original.</summary>
        public static Exception Wrap(Exception exception) =>
            new InvalidOperationException(exception.GetType().FullName + ": " + exception.Message + Environment.NewLine + exception);

        private static IEnumerable<Exception> Chain(Exception exception)
        {
            var pending = new Stack<Exception>();
            if (exception != null) pending.Push(exception);
            while (pending.Count > 0)
            {
                Exception e = pending.Pop();
                yield return e;
                if (e is AggregateException aggregate)
                    foreach (Exception inner in aggregate.InnerExceptions) pending.Push(inner);
                else if (e is ReflectionTypeLoadException typeLoad)
                    foreach (Exception inner in typeLoad.LoaderExceptions) { if (inner != null) pending.Push(inner); }
                else if (e.InnerException != null)
                    pending.Push(e.InnerException);
            }
        }
    }
}
