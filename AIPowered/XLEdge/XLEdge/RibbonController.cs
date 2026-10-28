using System;
using System.Collections.Generic;
using System.Reflection;
using XLEdge.Contracts;
using XLEdge.Hosting;

namespace XLEdge
{
    /// <summary>
    /// The ribbon and the task panes as the core sees them (IRibbonController). Controls are found by their AddinModule field
    /// name (RibbonControlIds) through reflection, because the core has no Add-in Express reference. Every member catches and
    /// logs its own errors: nothing may escape back across the AppDomain boundary.
    /// </summary>
    internal sealed class RibbonController : MarshalByRefObject, IRibbonController
    {
        private readonly object _controlOwner;
        private readonly IShellCommands _shell;
        private readonly ILogger _logger;
        private readonly Dictionary<string, object> _controls = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        /// <param name="controlOwner">The object whose fields are the ribbon controls (AddinModule).</param>
        public RibbonController(object controlOwner, IShellCommands shell, ILogger logger)
        {
            _controlOwner = controlOwner ?? throw new ArgumentNullException(nameof(controlOwner));
            _shell = shell ?? throw new ArgumentNullException(nameof(shell));
            _logger = logger;
        }

        /// <summary>
        /// Runs a control's action as if it was clicked. A reload is queued, never run on this call: the core may be the caller,
        /// and unloading its AppDomain under its own stack would abort Excel's thread.
        /// </summary>
        public void ExecuteAction(string controlId) => Guard(nameof(ExecuteAction), () =>
        {
            if (controlId == RibbonControlIds.RibReload || controlId == RibbonControlIds.RibVersionHistory)
                _shell.Post(() => _shell.RunRibbonAction(controlId));
            else
                _shell.RunRibbonAction(controlId);
        });

        /// <summary>The ribbon states of the XLEdge add-in before the shell/core split (XLEdgeRibbonHelper), unchanged.</summary>
        public void SetState(string stateName) => Guard(nameof(SetState), () =>
        {
            switch (stateName)
            {
                case "LoggedOut":
                    foreach (string id in RibbonControlIds.LoggedOutDisabled) Set(id, "Enabled", false);
                    foreach (string id in RibbonControlIds.LoggedOutEnabled) Set(id, "Enabled", true);
                    Set(RibbonControlIds.RibEdgeLogin, "Visible", true);
                    Set(RibbonControlIds.RibEdgeLogout, "Visible", false);
                    Set(RibbonControlIds.RibEdgeDebug, "Pressed", false);
                    Set(RibbonControlIds.RibEdgeIncludeOutputData, "Pressed", false);
                    Set(RibbonControlIds.RibEdgeLogin, "Caption", "Login");
                    _shell.DebugModeChanged(false);
                    break;
                case "LoggedIn":
                    foreach (string id in RibbonControlIds.LoggedOutDisabled) Set(id, "Enabled", true);
                    Set(RibbonControlIds.RibEdgeLogin, "Visible", false);
                    Set(RibbonControlIds.RibEdgeLogout, "Visible", true);
                    Set(RibbonControlIds.RibEdgeIncludeOutputData, "Pressed", false);
                    break;
                case "NoXLEdgePermission":
                    Set(RibbonControlIds.RibEdgeIncludeOutputData, "Pressed", false);
                    Set(RibbonControlIds.RibEdgeDebug, "Pressed", false);
                    foreach (string id in RibbonControlIds.NoPermissionDisabled) Set(id, "Enabled", false);
                    _shell.DebugModeChanged(false);
                    break;
                default:
                    _logger?.LogWarn($"RibbonController.SetState: unknown ribbon state '{stateName}'.");
                    return;
            }
            _shell.InvalidateRibbon();
        });

        public void SetControlEnabled(string controlId, bool enabled) => Guard(nameof(SetControlEnabled), () => Set(controlId, "Enabled", enabled));
        public void SetControlVisible(string controlId, bool visible) => Guard(nameof(SetControlVisible), () => Set(controlId, "Visible", visible));
        public void SetControlPressed(string controlId, bool pressed) => Guard(nameof(SetControlPressed), () => Set(controlId, "Pressed", pressed));
        public void SetControlLabel(string controlId, string label) => Guard(nameof(SetControlLabel), () => Set(controlId, "Caption", label));

        public bool GetControlPressed(string controlId) => Guard(nameof(GetControlPressed), () => Get(controlId, "Pressed") is bool b && b);
        public bool GetControlVisible(string controlId) => Guard(nameof(GetControlVisible), () => Get(controlId, "Visible") is bool b && b);
        public string GetControlLabel(string controlId) => Guard(nameof(GetControlLabel), () => Get(controlId, "Caption") as string) ?? string.Empty;

        public void EnableControls(string[] controlIds) => Guard(nameof(EnableControls), () => { foreach (string id in controlIds ?? new string[0]) Set(id, "Enabled", true); });
        public void DisableControls(string[] controlIds) => Guard(nameof(DisableControls), () => { foreach (string id in controlIds ?? new string[0]) Set(id, "Enabled", false); });

        /// <summary>Invalidates one control by its RibbonX Id (Office ignores the field name).</summary>
        public void Invalidate(string controlId) => Guard(nameof(Invalidate), () =>
        {
            if (Get(controlId, "Id") is string ribbonXId && !string.IsNullOrEmpty(ribbonXId)) _shell.InvalidateControl(ribbonXId);
        });

        public void InvalidateAll() => Guard(nameof(InvalidateAll), () => _shell.InvalidateRibbon());

        public void ShowTaskPane() => Guard(nameof(ShowTaskPane), () => _shell.ShowTaskPane());
        public void ToggleTaskPane() => Guard(nameof(ToggleTaskPane), () => _shell.ToggleTaskPane());
        public void HideAllTaskPanes() => Guard(nameof(HideAllTaskPanes), () => _shell.HideAllTaskPanes());
        public bool IsTaskPaneVisible() => Guard(nameof(IsTaskPaneVisible), () => _shell.IsTaskPaneVisible());
        public void SetTaskPaneCaption(string caption) => Guard(nameof(SetTaskPaneCaption), () => _shell.SetTaskPaneCaption(caption ?? string.Empty));

        public override object InitializeLifetimeService() => null;

        // ---- reflection over the ribbon control fields -----------------------------------

        private object Control(string controlId)
        {
            if (string.IsNullOrEmpty(controlId)) return null;
            if (_controls.TryGetValue(controlId, out object cached)) return cached;
            const BindingFlags Flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            Type type = _controlOwner.GetType();
            object control = type.GetProperty(controlId, Flags)?.GetValue(_controlOwner) ?? type.GetField(controlId, Flags)?.GetValue(_controlOwner);
            if (control == null)
            {
                _logger?.LogWarn($"RibbonController: ribbon control '{controlId}' does not exist; the call is ignored.");
                return null;
            }
            _controls[controlId] = control;
            return control;
        }

        private void Set(string controlId, string property, object value)
        {
            object control = Control(controlId);
            if (control == null) return;
            PropertyInfo info = control.GetType().GetProperty(property);
            if (info == null || !info.CanWrite)
            {
                _logger?.LogWarn($"RibbonController: '{controlId}' has no settable {property}; the call is ignored.");
                return;
            }
            info.SetValue(control, value);
        }

        private object Get(string controlId, string property)
        {
            object control = Control(controlId);
            return control?.GetType().GetProperty(property)?.GetValue(control);
        }

        private void Guard(string member, Action action)
        {
            try { action(); }
            catch (Exception ex) { _logger?.LogError($"RibbonController.{member} failed.", ex.ToString()); }
        }

        private T Guard<T>(string member, Func<T> func)
        {
            try { return func(); }
            catch (Exception ex)
            {
                _logger?.LogError($"RibbonController.{member} failed.", ex.ToString());
                return default(T);
            }
        }
    }
}
