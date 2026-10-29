using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization.Formatters.Binary;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;

namespace XLEdge.Tests
{
    /// <summary>
    /// Pins the exact shell/core contract. A failing test here means the contract changed:
    /// after the first MSI that needs ContractInfo.Version + 1 and a new MSI (spec §3).
    /// </summary>
    [TestClass]
    public class ContractSurfaceTests
    {
        private static readonly Assembly ContractsAssembly = typeof(ContractInfo).Assembly;

        [TestMethod]
        public void ExportedTypes_AreExactlyTheContractSet()
        {
            var actual = ContractsAssembly.GetExportedTypes().Select(t => t.FullName).OrderBy(n => n, StringComparer.Ordinal);
            var expected = new[]
            {
                "XLEdge.Contracts.ContractInfo", "XLEdge.Contracts.ILogger", "XLEdge.Contracts.IPathProvider",
                "XLEdge.Contracts.IRibbonController", "XLEdge.Contracts.IXLEdgeAddin", "XLEdge.Contracts.IXLEdgeContext",
                "XLEdge.Contracts.LoginInfo", "XLEdge.Contracts.VersionInfo",
            };
            CollectionAssert.AreEqual(expected, actual.ToArray());
        }

        [TestMethod]
        public void References_AreFrameworkOnly()
        {
            var names = ContractsAssembly.GetReferencedAssemblies().Select(a => a.Name).ToArray();
            CollectionAssert.AreEquivalent(new[] { "mscorlib" }, names);
        }

        [TestMethod]
        public void ContractInfo_IsStaticWithTheCurrentVersion()
        {
            Assert.IsTrue(typeof(ContractInfo).IsAbstract && typeof(ContractInfo).IsSealed, "ContractInfo must be a static class");
            Assert.AreEqual(4, ContractInfo.Version, "4: IRibbonController, the pane of an Excel window (Stage 2e)");
            AssertMembers(typeof(ContractInfo), "field const int Version");
        }

        [TestMethod]
        public void VersionInfo_Members()
        {
            Assert.IsTrue(typeof(VersionInfo).IsClass);
            AssertMembers(typeof(VersionInfo),
                "ctor()",
                "string Version {get;set;}", "string ReleaseDate {get;set;}", "string DownloadUrl {get;set;}",
                "string Checksum {get;set;}", "string Notes {get;set;}", "bool Mandatory {get;set;}",
                "string FolderName {get;set;}", "string FileName {get;set;}", "int ContractVersion {get;set;}");
        }

        [TestMethod]
        public void LoginInfo_Members()
        {
            Assert.IsTrue(typeof(LoginInfo).IsClass);
            AssertMembers(typeof(LoginInfo),
                "ctor()",
                "string LoginUrl {get;set;}", "string LoginToken {get;set;}", "string LoginUserName {get;set;}", "bool IsLoggedIn {get;set;}");
        }

        [TestMethod]
        public void VersionInfo_RoundTripsByValue()
        {
            var original = new VersionInfo
            {
                Version = "11.1.2.0", ReleaseDate = "2026-10-28T10:00:00", DownloadUrl = "https://example/x.zip", Checksum = "ab12",
                Notes = "first", Mandatory = true, FolderName = "v11.1.2.0_2026-10-28T10-00-00",
                FileName = "v11.1.2.0_2026-10-28T10-00-00.zip", ContractVersion = 1,
            };
            var copy = RoundTrip(original);
            foreach (var p in typeof(VersionInfo).GetProperties())
                Assert.AreEqual(p.GetValue(original), p.GetValue(copy), p.Name);
        }

        [TestMethod]
        public void LoginInfo_RoundTripsByValue()
        {
            var original = new LoginInfo { LoginUrl = "https://orbit.example/reporting", LoginToken = "tok", LoginUserName = "qa.user", IsLoggedIn = true };
            var copy = RoundTrip(original);
            foreach (var p in typeof(LoginInfo).GetProperties())
                Assert.AreEqual(p.GetValue(original), p.GetValue(copy), p.Name);
        }

        [TestMethod]
        public void ILogger_Members()
        {
            Assert.IsTrue(typeof(ILogger).IsInterface);
            AssertMembers(typeof(ILogger),
                "void LogInfo(string message)", "void LogWarn(string message)",
                "void LogError(string message, string exceptionDetails = null)", "void LogDebug(string message)",
                "void LogException(string exceptionDetails, string context = \"\")", "void LogRawJson(string context, string rawJson)",
                "void FlushDebugLogs(string section = \"Buffered Logs\")",
                "void BeginReleaseSession(string version, string releaseDate)", "void EndReleaseSession(string version, string releaseDate)",
                "IDisposable BeginLogScope(string scopeName)");
        }

        [TestMethod]
        public void IPathProvider_Members()
        {
            Assert.IsTrue(typeof(IPathProvider).IsInterface);
            AssertMembers(typeof(IPathProvider),
                "string Root {get;}", "string Logs {get;}", "string Temp {get;}", "string BrowserDataPath {get;}", "string UrlsFile {get;}",
                "string VersionsPath {get;}", "string SharedDependenciesPath {get;}", "string ManifestFile {get;}",
                "string ManifestDirectory {get;}", "string ReleaseHistoryFile {get;}",
                "string LatestVersion {get;}", "string LatestReleaseDate {get;}",
                "void Ensure()", "void Refresh()");
        }

        [TestMethod]
        public void IRibbonController_Members()
        {
            Assert.IsTrue(typeof(IRibbonController).IsInterface);
            AssertMembers(typeof(IRibbonController),
                "void ExecuteAction(string controlId)", "void SetState(string stateName)",
                "void SetControlEnabled(string controlId, bool enabled)", "void SetControlVisible(string controlId, bool visible)",
                "void SetControlPressed(string controlId, bool pressed)", "bool GetControlPressed(string controlId)",
                "bool GetControlVisible(string controlId)", "void SetControlLabel(string controlId, string label)",
                "string GetControlLabel(string controlId)", "void EnableControls(string[] controlIds)",
                "void DisableControls(string[] controlIds)", "void Invalidate(string controlId)", "void InvalidateAll()",
                "void ShowTaskPane()", "void ToggleTaskPane()", "void HideAllTaskPanes()", "bool IsTaskPaneVisible()",
                "void SetTaskPaneCaption(string caption)",
                "IntPtr GetActiveTaskPaneWindow()", "bool GetTaskPaneVisible(IntPtr excelWindow)",
                "void SetTaskPaneVisible(IntPtr excelWindow, bool visible)", "void SetTaskPaneCaption(IntPtr excelWindow, string caption)",
                "int GetTaskPaneWidth(IntPtr excelWindow)", "void SetTaskPaneWidth(IntPtr excelWindow, int width)");
        }

        [TestMethod]
        public void IXLEdgeContext_Members()
        {
            Assert.IsTrue(typeof(IXLEdgeContext).IsInterface);
            AssertMembers(typeof(IXLEdgeContext),
                "object ExcelApp {get;}", "ILogger Logger {get;}", "IPathProvider Paths {get;}",
                "IRibbonController RibbonController {get;}", "IntPtr ExcelHandle {get;set;}", "bool DebugMode {get;set;}",
                "string Version {get;set;}", "string ReleaseDate {get;set;}", "string ActiveFolderName {get;set;}",
                "string ShellVersion {get;}", "void SetRibbonController(IRibbonController controller)",
                "object GetGLSenseAddinInstance()", "IntPtr GetExcelDispatch()",
                "string GetSessionValue(string key)", "void SetSessionValue(string key, string value)");
        }

        [TestMethod]
        public void IXLEdgeAddin_Members()
        {
            Assert.IsTrue(typeof(IXLEdgeAddin).IsInterface);
            AssertMembers(typeof(IXLEdgeAddin),
                "void Initialize(IXLEdgeContext context)", "void OnRibbonAction(string action, object parameter)",
                "bool OnExcelEvent(string eventName, object[] args)", "void Shutdown()",
                "IntPtr CreateTaskPaneContent(IntPtr excelWindowHandle)", "void CloseTaskPaneContent(IntPtr excelWindowHandle)",
                "void OnTaskPaneShown(IntPtr excelWindowHandle)",
                "void OnLoginFromGLSense(string instanceName, string loginUrl, string token, string userName, bool hasXLEdgePermission)",
                "void OnLogoutFromGLSense()", "LoginInfo GetLoginInfo()", "string[] GetOpenWindowTitles()", "bool IsBusy()");
        }

        // ---- helpers -------------------------------------------------------------

        private static T RoundTrip<T>(T value)
        {
            Assert.IsTrue(typeof(T).IsSerializable, typeof(T).Name + " must be [Serializable]");
            var formatter = new BinaryFormatter();
            using (var stream = new MemoryStream())
            {
#pragma warning disable SYSLIB0011 // BinaryFormatter is what .NET Remoting uses for by-value objects
                formatter.Serialize(stream, value);
                stream.Position = 0;
                return (T)formatter.Deserialize(stream);
#pragma warning restore SYSLIB0011
            }
        }

        private static void AssertMembers(Type type, params string[] expected)
        {
            var actual = Describe(type).ToList();
            var missing = expected.Where(e => !actual.Contains(e, StringComparer.Ordinal)).ToList();
            var extra = actual.Where(a => !expected.Contains(a, StringComparer.Ordinal)).ToList();
            if (missing.Count > 0 || extra.Count > 0)
                Assert.Fail($"{type.Name}: missing [{string.Join("; ", missing)}] unexpected [{string.Join("; ", extra)}]");
        }

        /// <summary>Every public member of the type as a signature string (case-sensitive), including base interfaces, events and fields.</summary>
        private static IEnumerable<string> Describe(Type type)
        {
            const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;
            foreach (var i in type.GetInterfaces()) yield return "implements " + i.Name;
            if (type.BaseType != null && type.BaseType != typeof(object)) yield return "extends " + type.BaseType.Name;
            foreach (var c in type.GetConstructors(BindingFlags.Public | BindingFlags.Instance))
                yield return "ctor(" + string.Join(", ", c.GetParameters().Select(Param)) + ")";
            foreach (var f in type.GetFields(flags))
                yield return "field " + (f.IsLiteral ? "const " : f.IsStatic ? "static " : "") + Name(f.FieldType) + " " + f.Name;
            foreach (var e in type.GetEvents(flags)) yield return "event " + Name(e.EventHandlerType) + " " + e.Name;
            foreach (var p in type.GetProperties(flags))
                yield return $"{Name(p.PropertyType)} {p.Name} {{{(p.GetGetMethod() != null ? "get;" : "")}{(p.GetSetMethod() != null ? "set;" : "")}}}";
            foreach (var m in type.GetMethods(flags).Where(m => !m.IsSpecialName))
                yield return $"{Name(m.ReturnType)} {m.Name}({string.Join(", ", m.GetParameters().Select(Param))})";
        }

        private static string Param(ParameterInfo p)
        {
            var s = (p.GetCustomAttribute<CallerMemberNameAttribute>() != null ? "[CallerMemberName] " : "") + Name(p.ParameterType) + " " + p.Name;
            if (p.HasDefaultValue)
                s += p.DefaultValue == null ? " = null" : p.DefaultValue is string str ? $" = \"{str}\"" : " = " + p.DefaultValue;
            return s;
        }

        private static string Name(Type t)
        {
            if (t.IsArray) return Name(t.GetElementType()) + "[]";
            if (t.IsGenericType)
                return t.Name.Substring(0, t.Name.IndexOf('`')) + "<" + string.Join(",", t.GetGenericArguments().Select(Name)) + ">";
            if (t.Namespace != "XLEdge.Contracts" && t.Namespace != "System")
                Assert.Fail("Contract signature uses a type outside System / XLEdge.Contracts: " + t.FullName);
            switch (t.FullName)
            {
                case "System.String": return "string";
                case "System.Boolean": return "bool";
                case "System.Void": return "void";
                case "System.Object": return "object";
                case "System.Int32": return "int";
                default: return t.Name;
            }
        }
    }
}
