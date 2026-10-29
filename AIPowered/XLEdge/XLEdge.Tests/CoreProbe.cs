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

        /// <summary>The golden cases (GoldenCases) evaluated against the core in this domain.</summary>
        public string[] RunGoldenCases() => GoldenCases.Run(CoreType("XLEdge.Addin.Core.AddinEntry").Assembly);

        /// <summary>Whether a static member is non-null (for members whose value cannot cross the domain boundary).</summary>
        public bool HasStatic(string typeName, string member) => CallStatic(typeName, member) != null;

        private readonly System.Collections.Generic.List<object> _instances = new System.Collections.Generic.List<object>();

        /// <summary>Creates an instance of a core type in this domain; returns a handle for <see cref="CallOnInstance"/>.</summary>
        public int CreateInstance(string typeName, params object[] args)
        {
            // By parameter count and Invoke: the binder behind Activator.CreateInstance does not see a remoting proxy as the
            // interface it implements, Invoke checks it at the call.
            ConstructorInfo ctor = CoreType(typeName).GetConstructors(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                                       .FirstOrDefault(c => c.GetParameters().Length == (args?.Length ?? 0))
                                   ?? throw new MissingMethodException(typeName, ".ctor");
            object instance = Unwrap(() => ctor.Invoke(args));
            _instances.Add(instance);
            return _instances.Count - 1;
        }

        /// <summary>Calls an instance method (matched by name and argument count) of an object made by <see cref="CreateInstance"/>.</summary>
        public object CallOnInstance(int handle, string method, params object[] args)
        {
            object instance = _instances[handle];
            MethodInfo info = instance.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                                  .FirstOrDefault(m => m.Name == method && m.GetParameters().Length == (args?.Length ?? 0))
                              ?? throw new MissingMethodException(instance.GetType().FullName, method);
            return Unwrap(() => info.Invoke(instance, args));
        }

        /// <summary>WPF in this domain: creates the Application through WpfAppManager and runs a call through UiDispatcher.</summary>
        public string BootWpfAndRunOnTheDispatcher()
        {
            CallStatic("XLEdge.Utilities.WpfAppManager", "EnsureApplication");
            CallStatic("XLEdge.Utilities.WpfAppManager", "EnsureApplication");
            int sameThread = -1;
            Action action = () => sameThread = System.Threading.Thread.CurrentThread.ManagedThreadId;
            CallStatic("XLEdge.Utilities.UiDispatcher", "Run", action);
            bool hasApp = System.Windows.Application.Current != null;
            return $"app={hasApp} thread={(sameThread == System.Threading.Thread.CurrentThread.ManagedThreadId ? "caller" : sameThread.ToString())}";
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

        private object _kept;

        /// <summary>Holds the core's current Excel object the way ported code does (var app = XLApp.App).</summary>
        public void KeepExcel() => _kept = CallStatic("XLEdge.Addin.Core.Infrastructure.CoreServices", "ExcelApp");

        /// <summary>Calls a property or method on the object <see cref="KeepExcel"/> holds.</summary>
        public object InvokeOnKept(string member) =>
            _kept.GetType().InvokeMember(member, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, _kept, null);

        /// <summary>XLEdge 11.1.2's AddinBeginShutdown sequence (AddinModule.cs:1827-1831), then the core's own Shutdown.</summary>
        public void RunTodaysShutdownSequence()
        {
            object cached = OnSingleton("XLEdge.XLEdgeAppState", "Instance", "ExcelAppUnsafe");
            if (cached != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cached);
            object state = CallStatic("XLEdge.XLEdgeAppState", "Instance");
            state.GetType().GetMethod("ClearExcelApplicationUnsafe", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public).Invoke(state, null);
            CallStatic("XLEdge.Addin.Core.Infrastructure.CoreServices", "Shutdown");
        }

        /// <summary>The real entry point: after Initialize, XLEdgeRibbonHelper.Current talks to the context's ribbon controller.</summary>
        public void InitializeTheEntryAndDisableThroughTheRibbonHelper(IXLEdgeContext context, string controlId)
        {
            var entry = (IXLEdgeAddin)Activator.CreateInstance(CoreType("XLEdge.Addin.Core.AddinEntry"));
            entry.Initialize(context);
            object helper = CallStatic("XLEdge.Helpers.XLEdgeRibbonHelper", "Current");
            helper.GetType().GetMethod("SetControlEnabled").Invoke(helper, new object[] { controlId, false });
            entry.Shutdown();
        }

        private IXLEdgeAddin _entry;

        /// <summary>The real entry point with WPF booted in this domain (as the ported windows will do); see <see cref="ShutDownTheEntry"/>.</summary>
        public void InitializeTheEntryWithWpf(IXLEdgeContext context)
        {
            _entry = (IXLEdgeAddin)Activator.CreateInstance(CoreType("XLEdge.Addin.Core.AddinEntry"));
            _entry.Initialize(context);
            CallStatic("XLEdge.Utilities.WpfAppManager", "EnsureApplication");
        }

        public void ShutDownTheEntry() => _entry?.Shutdown();

        public string[] EntryOpenWindowTitles() => _entry.GetOpenWindowTitles();

        /// <summary>XLEdgeUserPreferences from "0"/"1" flags (declaration order), applied as Options does: Apply (session) or Save.</summary>
        public void SetOptions(string flags, bool save)
        {
            Type type = CoreType("XLEdge.Models.XLEdgeUserPreferences");
            object preferences = Activator.CreateInstance(type);
            string[] names = { "ParameterValues", "ScheduledOutputs", "RefreshSync", "ChangeSheetName", "CalendarCtrlDisplay", "SegmentSelectionWindowDisplay", "PreserveFormats" };
            for (int i = 0; i < names.Length; i++) type.GetProperty(names[i]).SetValue(preferences, flags[i] == '1');
            CallOnSingleton("XLEdge.Helpers.XLEdgePreferencesManager", "Instance", save ? "Save" : "ApplyRuntime", preferences);
        }

        /// <summary>The options in force (XLEdgeAppState) as the same seven "0"/"1" flags.</summary>
        public string OptionsInForce()
        {
            string[] state = { "ParamDataSameSheet", "SchOutputsToSameSheet", "RefreshSync", "AllowSheetNameChanges", "ShowCalendarControl", "ShowSegmentSelectionWindow", "PreserveFormats" };
            return string.Concat(state.Select(p => (bool)OnSingleton("XLEdge.XLEdgeAppState", "Instance", p) ? "1" : "0"));
        }

        /// <summary>This domain's ServicePointManager after the core's ApiHelper type is initialised (its static constructor).</summary>
        public string ServicePointSettingsAfterApiHelper()
        {
            System.Runtime.CompilerServices.RuntimeHelpers.RunClassConstructor(CoreType("XLEdge.Helpers.ApiHelper").TypeHandle);
            return $"limit={System.Net.ServicePointManager.DefaultConnectionLimit} expect100={System.Net.ServicePointManager.Expect100Continue} " +
                   $"idle={System.Net.ServicePointManager.MaxServicePointIdleTime} protocols={System.Net.ServicePointManager.SecurityProtocol}";
        }

        /// <summary>ApiHelper.ServerAPI from this domain, with a login token set: "ok:&lt;body&gt;" or "throws &lt;type&gt;: &lt;message&gt;".</summary>
        public string CallServerApi(string url, string postData, string token)
        {
            OnSingleton("XLEdge.XLEdgeAppState", "Instance", "LoginToken", token);
            MethodInfo serverApi = CoreType("XLEdge.Helpers.ApiHelper").GetMethod("ServerAPI");
            try
            {
                var task = (System.Threading.Tasks.Task<string>)serverApi.Invoke(null, new object[] { url, "application/json", postData, "POST", default(System.Threading.CancellationToken) });
                return "ok:" + task.GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Exception e = ex is TargetInvocationException tie && tie.InnerException != null ? tie.InnerException : ex;
                return "throws " + e.GetType().FullName + ": " + e.Message;
            }
        }

        /// <summary>
        /// Opens a core window (non-modal) on this STA thread and lets WPF load and lay it out; a user control is hosted in a
        /// plain window. Returns "loaded=..., size=WxH" - a XAML or resource failure throws instead.
        /// </summary>
        public string OpenWindow(string typeName, params object[] args)
        {
            Type type = CoreType(typeName);
            ConstructorInfo ctor = type.GetConstructors().First(c => c.GetParameters().Length == (args?.Length ?? 0));
            object created = Unwrap(() => ctor.Invoke(args));
            var window = created as System.Windows.Window ?? new System.Windows.Window { Title = typeName, Content = created, Width = 400, Height = 300 };
            window.ShowActivated = false;
            string name = typeName;
            _opened.Add(name);
            window.Closed += (s, e) => _closed.Add(name);
            Unwrap(() => { window.Show(); return null; });
            PumpWpf();
            return $"loaded={window.IsLoaded}, size={(int)window.ActualWidth}x{(int)window.ActualHeight}";
        }

        private readonly System.Collections.Generic.List<string> _opened = new System.Collections.Generic.List<string>();
        private readonly System.Collections.Generic.List<string> _closed = new System.Collections.Generic.List<string>();

        /// <summary>Windows opened by <see cref="OpenWindow"/> that never raised Closed (a window that vetoed its closing).</summary>
        public string[] WindowsNotClosed() => _opened.Except(_closed).ToArray();

        /// <summary>Source URIs of every merged resource dictionary of this domain's Application (nested ones included).</summary>
        public string[] ResourceDictionarySources()
        {
            var sources = new System.Collections.Generic.List<string>();
            void Walk(System.Windows.ResourceDictionary dictionary)
            {
                if (dictionary.Source != null) sources.Add(dictionary.Source.OriginalString);
                foreach (var merged in dictionary.MergedDictionaries) Walk(merged);
            }
            if (System.Windows.Application.Current != null) Walk(System.Windows.Application.Current.Resources);
            return sources.ToArray();
        }

        /// <summary>IsBusy as asked from inside a modal core dialog (the reload's check, if Excel's message loop ran it then).</summary>
        public bool IsBusyInsideAModalDialog()
        {
            // A real XLEdge dialog (a bare Window clashes with the theme's window style under WPF-UI).
            var dialog = (System.Windows.Window)Activator.CreateInstance(CoreType("XLEdge.Views.XLEdgeLoginDetails"));
            bool busy = false;
            dialog.Dispatcher.BeginInvoke(new Action(() => { busy = _entry.IsBusy(); dialog.Close(); }), System.Windows.Threading.DispatcherPriority.Background);
            dialog.ShowDialog();
            return busy;
        }

        public bool EntryIsBusy() => _entry.IsBusy();

        /// <summary>Lets the dispatcher of this thread (in this domain) run everything queued down to Background.</summary>
        private static void PumpWpf()
        {
            for (int i = 0; i < 5; i++)
                System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() => { }));
        }

        /// <summary>The real entry point, initialized (see <see cref="ShutDownTheEntry"/>).</summary>
        public void InitializeTheEntry(IXLEdgeContext context)
        {
            _entry = (IXLEdgeAddin)Activator.CreateInstance(CoreType("XLEdge.Addin.Core.AddinEntry"));
            _entry.Initialize(context);
        }

        /// <summary>Calls a method of a singleton reached through a static member (e.g. XLEdgePreferencesManager.Instance.Initialize()).</summary>
        public object CallOnSingleton(string typeName, string instanceMember, string method, params object[] args)
        {
            object instance = CallStatic(typeName, instanceMember);
            MethodInfo info = instance.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                                  .FirstOrDefault(m => m.Name == method && m.GetParameters().Length == (args?.Length ?? 0))
                              ?? throw new MissingMethodException(typeName, method);
            return Unwrap(() => info.Invoke(instance, args));
        }

        /// <summary>The real entry point: Initialize, make the core take Excel, then Shutdown.</summary>
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
