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

        // Contract 4: the pane of one Excel window, keyed by that window's handle (the key CreateTaskPaneContent receives).
        // Zero when no Excel window's pane is active; members given an Excel window without a pane do nothing (or return false / 0).
        System.IntPtr GetActiveTaskPaneWindow();
        bool GetTaskPaneVisible(System.IntPtr excelWindow);
        /// <summary>Showing a window's pane creates it when it does not exist yet (only for the active Excel window).</summary>
        void SetTaskPaneVisible(System.IntPtr excelWindow, bool visible);
        void SetTaskPaneCaption(System.IntPtr excelWindow, string caption);
        int GetTaskPaneWidth(System.IntPtr excelWindow);
        void SetTaskPaneWidth(System.IntPtr excelWindow, int width);

        // Contract 6: the post-login add-in update check (login -> update check -> broadcast message).
        /// <summary>
        /// Asks the server (with the given login) for an XLEdge release newer than the loaded one, within <paramref name="timeoutSeconds"/>.
        /// Returns its display text ("version (release date)"), or null when there is none or the check failed (failures are only logged).
        /// Call it off Excel's thread: it waits for the server.
        /// </summary>
        string FindNewerAddinRelease(string loginUrl, string loginToken, int timeoutSeconds);
        /// <summary>The same answer from GLSense when it is installed too (GLSense judges its own release); null when it is
        /// missing, disabled, older, or its check failed. Call it off Excel's thread.</summary>
        string FindNewerPartnerAddinRelease(string loginUrl, string loginToken, int timeoutSeconds);
        /// <summary>Downloads the release the last <see cref="FindNewerAddinRelease"/> found, logs out and reloads into it (queued on
        /// Excel's thread). <paramref name="includePartner"/>: GLSense reloads into its newer release first, then XLEdge.</summary>
        void ReloadToNewerAddinRelease(bool includePartner);
    }
}
