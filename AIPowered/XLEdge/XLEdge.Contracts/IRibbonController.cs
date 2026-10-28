namespace XLEdge.Contracts
{
    /// <summary>
    /// The shell's ribbon and task panes, driven from the core. Controls are addressed by their
    /// ribbon control id string, because the core has no Add-in Express reference. Lists are
    /// string arrays because arrays always cross the AppDomain boundary by value.
    /// Owned by the shell and remoted to the core, so its implementation must be a MarshalByRefObject.
    /// </summary>
    public interface IRibbonController
    {
        /// <summary>Runs the action bound to a ribbon control, as if the user clicked it.</summary>
        void ExecuteAction(string controlId);

        /// <summary>Applies a named ribbon state (for example logged out / logged in).</summary>
        void SetState(string stateName);

        void SetControlEnabled(string controlId, bool enabled);
        void SetControlVisible(string controlId, bool visible);
        void SetControlPressed(string controlId, bool pressed);
        bool GetControlPressed(string controlId);
        bool GetControlVisible(string controlId);
        void SetControlLabel(string controlId, string label);
        string GetControlLabel(string controlId);
        void EnableControls(string[] controlIds);
        void DisableControls(string[] controlIds);

        void Invalidate(string controlId);
        void InvalidateAll();

        // Task pane visibility is shell state (ADX task pane instances), one pane per Excel window.
        // ShowTaskPane, ToggleTaskPane and IsTaskPaneVisible act on the active Excel window's pane.
        void ShowTaskPane();
        void ToggleTaskPane();
        void HideAllTaskPanes();
        bool IsTaskPaneVisible();

        /// <summary>Sets the title of the XLEdge task panes (shown above the pane content).</summary>
        void SetTaskPaneCaption(string caption);
    }
}
