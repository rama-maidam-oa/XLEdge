namespace XLEdge.Addin.Core.Pane
{
    /// <summary>What 11.1.2 code asks of the active Excel window's task pane (AddinModule.GetPaneInstance); Stage 2e implements it.</summary>
    public interface ICorePane
    {
        /// <summary>Gives the keyboard back to Excel after a report run (ReportGenerator.ReleaseKeyboardFocusFromTaskPaneAsync).</summary>
        void ReleaseFocusToExcel();
    }
}
