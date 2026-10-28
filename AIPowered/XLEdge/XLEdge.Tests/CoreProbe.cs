using System;
using System.IO;
using System.Linq;
using System.Reflection;
using XLEdge.Contracts;

namespace XLEdge.Tests
{
    /// <summary>
    /// Runs inside an AppDomain rooted at the core's build folder (as the loader creates it) and calls the core's internal
    /// types by reflection - the core is not strong-named, so this signed test assembly cannot reference it, and shipped
    /// code carries no test hooks. Only primitives, strings and contract types cross back.
    /// </summary>
    public sealed class CoreProbe : MarshalByRefObject
    {
        public override object InitializeLifetimeService() => null;

        private static Type CoreType(string name)
        {
            Assembly core = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.GetName().Name == "XLEdge.Addin.Core")
                            ?? Assembly.LoadFrom(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "XLEdge.Addin.Core.dll"));
            return core.GetType(name, throwOnError: true);
        }

        /// <summary>Calls a static member (method, or property getter when <paramref name="args"/> is null and no such method exists).</summary>
        public object CallStatic(string typeName, string member, params object[] args)
        {
            const BindingFlags Flags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
            Type type = CoreType(typeName);
            MethodInfo method = type.GetMethods(Flags).FirstOrDefault(m => m.Name == member && m.GetParameters().Length == (args?.Length ?? 0));
            if (method != null) return Unwrap(() => method.Invoke(null, args));
            PropertyInfo property = type.GetProperty(member, Flags);
            if (property == null) throw new MissingMemberException(typeName, member);
            if (args != null && args.Length == 1) { Unwrap(() => { property.SetValue(null, args[0]); return null; }); return null; }
            return Unwrap(() => property.GetValue(null));
        }

        /// <summary>Reads or (with one argument) sets a property of a singleton reached through a static member (e.g. Instance).</summary>
        public object OnSingleton(string typeName, string instanceMember, string property, params object[] args)
        {
            object instance = CallStatic(typeName, instanceMember);
            PropertyInfo info = instance.GetType().GetProperty(property, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                                ?? throw new MissingMemberException(typeName, property);
            if (args != null && args.Length == 1) { Unwrap(() => { info.SetValue(instance, args[0]); return null; }); return null; }
            return Unwrap(() => info.GetValue(instance));
        }

        /// <summary>Calls LogUtility.LogException with an exception thrown (so it has a stack trace) inside this domain.</summary>
        public void LogThrownException(string message, string context)
        {
            try { throw new InvalidOperationException(message, new ArgumentException("inner " + message)); }
            catch (InvalidOperationException ex) { CallStatic("XLEdge.Utilities.LogUtility", "LogException", ex, context); }
        }

        /// <summary>Opens and closes a LogUtility.LogScope inside this domain, logging one debug line in it.</summary>
        public void RunInLogScope(string scopeName, string line)
        {
            Type scopeType = CoreType("XLEdge.Utilities.LogUtility+LogScope");
            using ((IDisposable)Activator.CreateInstance(scopeType, scopeName))
                CallStatic("XLEdge.Utilities.LogUtility", "LogDebug", line);
        }

        /// <summary>The core's view of its Excel object: is it a COM object of this domain, or a remoting proxy?</summary>
        public string DescribeExcel()
        {
            object excel = CallStatic("XLEdge.Addin.Core.Infrastructure.CoreServices", "ExcelApp");
            if (excel == null) return "null";
            if (System.Runtime.Remoting.RemotingServices.IsTransparentProxy(excel)) return "proxy";
            return System.Runtime.InteropServices.Marshal.IsComObject(excel) ? "com" : "other:" + excel.GetType().FullName;
        }

        /// <summary>Calls a method on the core's Excel object (late-bound, as ported code does through the interop types).</summary>
        public object InvokeOnExcel(string member, params object[] args)
        {
            object excel = CallStatic("XLEdge.Addin.Core.Infrastructure.CoreServices", "ExcelApp");
            return excel.GetType().InvokeMember(member, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, excel, args);
        }

        /// <summary>The real entry point in this domain: Initialize, make the core take Excel, then Shutdown.</summary>
        public void InitializeUseAndShutDownTheEntry(IXLEdgeContext context)
        {
            var entry = (IXLEdgeAddin)Activator.CreateInstance(CoreType("XLEdge.Addin.Core.AddinEntry"));
            entry.Initialize(context);
            if (DescribeExcel() != "com") throw new InvalidOperationException("the core did not take Excel: " + DescribeExcel());
            entry.Shutdown();
        }

        private static object Unwrap(Func<object> call)
        {
            try { return call(); }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                // Rethrown as text: a core-only exception type cannot cross back to the test domain.
                throw new InvalidOperationException(ex.InnerException.GetType().FullName + ": " + ex.InnerException.Message);
            }
        }

        /// <summary>An AppDomain like the loader's (rooted at the core build folder) with a probe in it.</summary>
        internal static (AppDomain Domain, CoreProbe Probe) Create(string name)
        {
            string coreFolder = ReleaseFolders.CoreBuildFolder;
            var domain = AppDomain.CreateDomain(name, null, new AppDomainSetup { ApplicationBase = coreFolder });
            var probe = (CoreProbe)domain.CreateInstanceFromAndUnwrap(typeof(CoreProbe).Assembly.Location, typeof(CoreProbe).FullName);
            return (domain, probe);
        }
    }
}
