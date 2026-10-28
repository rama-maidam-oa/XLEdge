using System;
using System.Collections.Generic;
using System.Reflection;
using AddinExpress.MSO;
using XLEdge.Contracts;

namespace XLEdge
{
    /// <summary>
    /// The ribbon and the task panes as the core sees them (IRibbonController). Controls are found by their AddinModule field
    /// name (RibbonControlIds) through reflection, because the core has no Add-in Express reference. Every member catches and
    /// logs its own errors: nothing may escape back across the AppDomain boundary.
    /// </summary>
    internal sealed class RibbonController : MarshalByRefObject, IRibbonController
    {
        private const int PaneWidth = 550;
        private readonly AddinModule _module;
        private readonly IRibbonUI _ribbon;
        private readonly ILogger _logger;
        private readonly Dictionary<string, object> _controls = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        public RibbonController(AddinModule module, IRibbonUI ribbon, ILogger logger)
        {
            _module = module ?? throw new ArgumentNullException(nameof(module));
            _ribbon = ribbon;
            _logger = logger;
        }

        public void ExecuteAction(string controlId) => Guard(nameof(ExecuteAction), () => _module.RunRibbonAction(controlId));

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
                    break;
                case "LoggedIn":
                    foreach (string id in RibbonControlIds.LoggedOutDisabled) Set(id, "Enabled", true);
                    Set(RibbonControlIds.RibEdgeLogin, "Visible", false);
                    Set(RibbonControlIds.RibEdgeLogout, "Visible", true);
                    break;
                case "NoXLEdgePermission":
                    foreach (string id in RibbonControlIds.AllFeatureControls) Set(id, "Enabled", false);
                    break;
                default:
                    _logger?.LogWarn($"RibbonController.SetState: unknown ribbon state '{stateName}'.");
                    return;
            }
            _ribbon?.Invalidate();
        });

        public void SetControlEnabled(string controlId, bool enabled) => Guard(nameof(SetControlEnabled), () => Set(controlId, "Enabled", enabled));
        public void SetControlVisible(string controlId, bool visible) => Guard(nameof(SetControlVisible), () => Set(controlId, "Visible", visible));
        public void SetControlPressed(string controlId, bool pressed) => Guard(nameof(SetControlPressed), () => Set(controlId, "Pressed", pressed));
        public void SetControlLabel(string controlId, string label) => Guard(nameof(SetControlLabel), () => Set(controlId, "Caption", label));

        public bool GetControlPressed(string controlId) => Guard(nameof(GetControlPressed), () => Get(controlId, "Pressed") is bool b && b);
        public bool GetControlVisible(string controlId) => Guard(nameof(GetControlVisible), () => Get(controlId, "Visible") is bool b && b);
        public string GetControlLabel(string controlId) => Guard(nameof(GetControlLabel), () => Get(controlId, "Caption") as string);

        public void EnableControls(string[] controlIds) => Guard(nameof(EnableControls), () => { foreach (string id in controlIds ?? new string[0]) Set(id, "Enabled", true); });
        public void DisableControls(string[] controlIds) => Guard(nameof(DisableControls), () => { foreach (string id in controlIds ?? new string[0]) Set(id, "Enabled", false); });

        /// <summary>Invalidates one control by its RibbonX Id (Office ignores the field name).</summary>
        public void Invalidate(string controlId) => Guard(nameof(Invalidate), () =>
        {
            if (Get(controlId, "Id") is string ribbonXId && !string.IsNullOrEmpty(ribbonXId)) _ribbon?.InvalidateControl(ribbonXId);
        });

        public void InvalidateAll() => Guard(nameof(InvalidateAll), () => _ribbon?.Invalidate());

        public void ShowTaskPane() => Guard(nameof(ShowTaskPane), () => _module.ShowTaskPane(PaneWidth));
        public void ToggleTaskPane() => Guard(nameof(ToggleTaskPane), () =>
        {
            if (_module.ActivePane?.Visible == true) _module.ActivePane.Visible = false;
            else _module.ShowTaskPane(PaneWidth);
        });
        public void HideAllTaskPanes() => Guard(nameof(HideAllTaskPanes), () => { foreach (XLEdgeTaskPane pane in _module.AllPanes()) pane.Visible = false; });
        public bool IsTaskPaneVisible() => Guard(nameof(IsTaskPaneVisible), () => _module.ActivePane?.Visible == true);
        public void SetTaskPaneCaption(string caption) => Guard(nameof(SetTaskPaneCaption), () => { foreach (XLEdgeTaskPane pane in _module.AllPanes()) pane.Text = caption ?? string.Empty; });

        public override object InitializeLifetimeService() => null;

        // ---- reflection over AddinModule's ribbon fields ---------------------------------

        private object Control(string controlId)
        {
            if (string.IsNullOrEmpty(controlId)) return null;
            if (_controls.TryGetValue(controlId, out object cached)) return cached;
            const BindingFlags Flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            Type type = _module.GetType();
            object control = type.GetProperty(controlId, Flags)?.GetValue(_module) ?? type.GetField(controlId, Flags)?.GetValue(_module);
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
            PropertyInfo info = control?.GetType().GetProperty(property);
            if (control == null) return;
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
