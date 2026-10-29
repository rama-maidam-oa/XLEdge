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

        // ---- Stage 2d: ribbon actions and Excel events as the shell forwards them ----------------------------------------------

        public void EntryRibbonAction(string action, object parameter) => _entry.OnRibbonAction(action, parameter);

        public bool EntryExcelEvent(string eventName, object[] args) => _entry.OnExcelEvent(eventName, args);

        /// <summary>
        /// An async-void method started the way AddinEntry starts 11.1.2's handlers (ExcelThreadContext.Run): which thread its
        /// continuation after an await runs on, and IsBusy while it is pending and once it finished.
        /// </summary>
        public string AsyncHandlerOnTheExcelThread()
        {
            MethodInfo run = CoreType("XLEdge.Addin.Core.Infrastructure.ExcelThreadContext").GetMethod("Run", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            int excelThread = System.Threading.Thread.CurrentThread.ManagedThreadId, resumedOn = -1;
            var gate = new System.Threading.Tasks.TaskCompletionSource<bool>();
            async void Handler()
            {
                await gate.Task;
                resumedOn = System.Threading.Thread.CurrentThread.ManagedThreadId;
            }
            Unwrap(() => run.Invoke(null, new object[] { (Action)Handler }));
            bool busyWhilePending = _entry.IsBusy();
            // Released from a pool thread: the continuation must still come back to this (Excel's) thread.
            System.Threading.Tasks.Task.Run(() => gate.SetResult(true));
            DateTime until = DateTime.UtcNow.AddSeconds(5);
            while (resumedOn < 0 && DateTime.UtcNow < until)
            {
                PumpWpf();
                System.Threading.Thread.Sleep(10);
            }
            return $"resumedOnExcelThread={resumedOn == excelThread} busyWhilePending={busyWhilePending} busyAfter={_entry.IsBusy()}";
        }

        /// <summary>
        /// A ribbon action, sent to the real entry point (AddinEntry.OnRibbonAction), whose handler - added to the module's handler
        /// table for this call - awaits: which thread it resumes on, and IsBusy while it is pending and once it ended.
        /// </summary>
        public string AwaitingHandlerThroughTheEntry()
        {
            const string Action = "Probe.AwaitingHandler";
            var table = (System.Collections.IDictionary)CoreType("XLEdge.AddinModule").GetField("RibbonHandlers", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
            int excelThread = System.Threading.Thread.CurrentThread.ManagedThreadId, resumedOn = -1;
            var gate = new System.Threading.Tasks.TaskCompletionSource<bool>();
            async void Handler(object module, bool pressed)
            {
                await gate.Task;
                resumedOn = System.Threading.Thread.CurrentThread.ManagedThreadId;
            }
            table[Action] = (Action<object, bool>)Handler;
            try
            {
                _entry.OnRibbonAction(Action, null);
                bool busyWhilePending = _entry.IsBusy();
                System.Threading.Tasks.Task.Run(() => gate.SetResult(true));
                WaitFor(() => resumedOn >= 0);
                return $"resumedOnExcelThread={resumedOn == excelThread} busyWhilePending={busyWhilePending} busyAfter={_entry.IsBusy()}";
            }
            finally
            {
                table.Remove(Action);
            }
        }

        /// <summary>An async-void handler started (ExcelThreadContext.Run) on a pool thread: does it still resume on Excel's thread?</summary>
        public string HandlerStartedOffExcelsThread()
        {
            MethodInfo run = CoreType("XLEdge.Addin.Core.Infrastructure.ExcelThreadContext").GetMethod("Run", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            int excelThread = System.Threading.Thread.CurrentThread.ManagedThreadId, resumedOn = -1;
            var gate = new System.Threading.Tasks.TaskCompletionSource<bool>();
            async void Handler()
            {
                await gate.Task;
                resumedOn = System.Threading.Thread.CurrentThread.ManagedThreadId;
            }
            System.Threading.Tasks.Task.Run(() => run.Invoke(null, new object[] { (Action)Handler })).Wait();
            System.Threading.Tasks.Task.Run(() => gate.SetResult(true));
            WaitFor(() => resumedOn >= 0);
            return $"resumedOnExcelThread={resumedOn == excelThread} busyAfter={_entry.IsBusy()}";
        }

        /// <summary>An async-void handler that throws (an IOException) after an await, started as AddinEntry starts handlers.</summary>
        public string HandlerThatThrowsAfterAnAwait()
        {
            MethodInfo run = CoreType("XLEdge.Addin.Core.Infrastructure.ExcelThreadContext").GetMethod("Run", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            var gate = new System.Threading.Tasks.TaskCompletionSource<bool>();
            bool reached = false;
            async void Handler()
            {
                await gate.Task;
                reached = true;
                throw new IOException("disk gone");
            }
            Unwrap(() => run.Invoke(null, new object[] { (Action)Handler }));
            System.Threading.Tasks.Task.Run(() => gate.SetResult(true));
            WaitFor(() => reached);
            PumpWpf();
            return $"reached={reached} busyAfter={_entry.IsBusy()}";
        }

        private static void WaitFor(Func<bool> done)
        {
            DateTime until = DateTime.UtcNow.AddSeconds(5);
            while (!done() && DateTime.UtcNow < until)
            {
                PumpWpf();
                System.Threading.Thread.Sleep(10);
            }
        }

        /// <summary>What the core reads from a forwarded event's arguments (AddinModule.Target), or "null".</summary>
        public string DescribeEventTarget(string eventName, object[] args)
        {
            MethodInfo target = CoreType("XLEdge.AddinModule").GetMethod("Target", BindingFlags.Static | BindingFlags.NonPublic);
            return Unwrap(() => target.Invoke(null, new object[] { eventName, args }))?.ToString() ?? "null";
        }

        /// <summary>
        /// The ribbon helper and a pane's ADX stand-ins (Visible, Text, Width), used from a pool thread as 11.1.2's code can after
        /// an await: the managed thread each shell call then arrived on, read from the ribbon (the test compares it with its own).
        /// </summary>
        public void ShellCallsFromAPoolThread(IntPtr excelWindow)
        {
            object helper = CallStatic("XLEdge.Helpers.XLEdgeRibbonHelper", "Current");
            object host = CoreType("XLEdge.Addin.Core.Infrastructure.CoreServices") != null ? _entry.GetType().GetField("_panes", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(_entry) : null;
            _entry.CreateTaskPaneContent(excelWindow);
            object pane = host.GetType().GetMethod("Pane").Invoke(host, new object[] { excelWindow });
            Exception failure = null;
            var work = System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    helper.GetType().GetMethod("SetControlCaption").Invoke(helper, new object[] { "RibSheetLabel", "from a pool thread" });
                    helper.GetType().GetMethod("SetControlEnabled").Invoke(helper, new object[] { "RibEdgeHelp", false });
                    pane.GetType().GetProperty("Visible", BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly).SetValue(pane, true);
                    pane.GetType().GetProperty("Text", BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly).SetValue(pane, "caption");
                    pane.GetType().GetProperty("Width", BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly).SetValue(pane, 600);
                }
                catch (Exception ex) { failure = ex; }
            });
            WaitFor(() => work.IsCompleted);
            if (failure != null) throw new InvalidOperationException(failure.ToString());
        }

        private object _module;

        /// <summary>The core's AddinModule with a deferred sheet delete queued and its timer running (as SheetBeforeDelete leaves it).</summary>
        public void QueueADeferredDelete()
        {
            _module = CallStatic("XLEdge.AddinModule", "CurrentInstance") ?? throw new InvalidOperationException("no AddinModule after Initialize");
            const BindingFlags Flags = BindingFlags.Instance | BindingFlags.NonPublic;
            ((System.Collections.Generic.List<string>)_module.GetType().GetField("_sheetsToDelete", Flags).GetValue(_module)).Add("P_Sheet1");
            ((System.Windows.Forms.Timer)_module.GetType().GetField("_deleteTimer", Flags).GetValue(_module)).Start();
        }

        /// <summary>The module queued by <see cref="QueueADeferredDelete"/>: its timer, its queue, and whether it is still the current one.</summary>
        public string DeferredDeleteState()
        {
            const BindingFlags Flags = BindingFlags.Instance | BindingFlags.NonPublic;
            var timer = (System.Windows.Forms.Timer)_module.GetType().GetField("_deleteTimer", Flags).GetValue(_module);
            var queue = (System.Collections.Generic.List<string>)_module.GetType().GetField("_sheetsToDelete", Flags).GetValue(_module);
            return $"timerEnabled={timer.Enabled} queued={queue.Count} current={CallStatic("XLEdge.AddinModule", "CurrentInstance") != null}";
        }

        private readonly System.Collections.Generic.List<string> _messages = new System.Collections.Generic.List<string>();
        private System.Windows.Threading.DispatcherTimer _answering;

        /// <summary>
        /// From now on every XLEdge message box the core shows is recorded and answered (Yes/OK) - they are modal - and the login
        /// windows (Server Configuration, About) are recorded as "window:name" and closed.
        /// </summary>
        public void AnswerMessages()
        {
            Type messageWindow = CoreType("XLEdge.Views.XLEdgeMessageWindow");
            _answering = new System.Windows.Threading.DispatcherTimer(System.Windows.Threading.DispatcherPriority.Background) { Interval = TimeSpan.FromMilliseconds(30) };
            _answering.Tick += (s, e) =>
            {
                if (System.Windows.Application.Current == null) return;
                foreach (System.Windows.Window window in System.Windows.Application.Current.Windows.Cast<System.Windows.Window>().ToList())
                {
                    if (!window.IsVisible) continue;
                    // The login windows (Server Configuration, About): recorded by name and closed.
                    if (window.GetType().Name == "XLEdgeServerConfiguration" || window.GetType().Name == "XLEdgeAbout")
                    {
                        _messages.Add("window:" + window.GetType().Name);
                        window.Close();
                        continue;
                    }
                    if (!messageWindow.IsInstanceOfType(window)) continue;
                    _messages.Add(((System.Windows.Controls.TextBlock)window.FindName("MsgText"))?.Text);
                    messageWindow.GetProperty("Result").SetValue(window, System.Windows.MessageBoxResult.Yes);
                    window.Close();
                }
            };
            _answering.Start();
        }

        public string[] MessagesShown()
        {
            PumpWpf();
            return _messages.ToArray();
        }

        public void StopAnsweringMessages() => _answering?.Stop();

        /// <summary>AddinModule.UpdateTabLabel on the active sheet of the core's Excel, as ProgressCoordinator calls it after a report run.</summary>
        public void UpdateTabLabelOfTheActiveSheet()
        {
            dynamic excel = CallStatic("XLEdge.Addin.Core.Infrastructure.CoreServices", "ExcelApp");
            object module = CallStatic("XLEdge.AddinModule", "CurrentInstance");
            object sheet = excel.ActiveSheet;
            Unwrap(() => module.GetType().GetMethod("UpdateTabLabel").Invoke(module, new[] { sheet }));
        }

        /// <summary>The sheet an Excel event naming (sheet, workbook) resolves to in the core: "workbook!its first table", or "null".</summary>
        public string ResolveEventSheet(string sheetName, string workbookName)
        {
            object module = CallStatic("XLEdge.AddinModule", "CurrentInstance");
            dynamic sheet = Unwrap(() => module.GetType().GetMethod("Sheet", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(module, new object[] { sheetName, workbookName }));
            return sheet == null ? "null" : (string)sheet.Parent.Name + "!" + (string)sheet.ListObjects[1].Name;
        }

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
