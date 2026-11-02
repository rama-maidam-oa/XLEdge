using System.Collections.ObjectModel;

namespace XLEdge
{
    /// <summary>
    /// Ribbon control ids used across the shell/core boundary: the AddinModule field names (RibbonController finds the controls
    /// by them). They are the same as in the XLEdge add-in before the shell/core split.
    /// </summary>
    internal static class RibbonControlIds
    {
        public const string RibEdgeLogin = "RibEdgeLogin";
        public const string RibEdgeLogout = "RibEdgeLogout";
        public const string RibEdgeBroadcast = "RibEdgeBroadcast";
        public const string RibLoginURL = "RibLoginURL";
        public const string RibEdgeDialogBoxLauncher = "RibEdgeDialogBoxLauncher";
        public const string RibEdgeRefresh = "RibEdgeRefresh";
        public const string RibEdgeRefreshAll = "RibEdgeRefreshAll";
        public const string RibEdgeParamRefresh = "RibEdgeParamRefresh";
        public const string RibEdgeParamRefreshBook = "RibEdgeParamRefreshBook";
        public const string RibEdgeShowHide = "RibEdgeShowHide";
        public const string RibEdgeDebug = "RibEdgeDebug";
        public const string RibEdgeIncludeOutputData = "RibEdgeIncludeOutputData";
        public const string RibEdgeOptions = "RibEdgeOptions";
        public const string RibControlSheet = "RibControlSheet";
        public const string RibEdgeAbout = "RibEdgeAbout";
        public const string RibEdgeHelp = "RibEdgeHelp";
        public const string RibSheetLabel = "RibSheetLabel";

        // Shell-only (hot reload); always enabled.
        public const string RibReload = "RibReload";
        public const string RibVersionHistory = "RibVersionHistory";

        /// <summary>Disabled while logged out (enabled when logged in).</summary>
        public static readonly ReadOnlyCollection<string> LoggedOutDisabled = new ReadOnlyCollection<string>(new[]
        {
            RibEdgeDialogBoxLauncher, RibEdgeRefresh, RibEdgeRefreshAll, RibEdgeParamRefresh, RibEdgeShowHide,
            RibEdgeIncludeOutputData, RibEdgeOptions, RibEdgeHelp, RibEdgeBroadcast,
        });

        /// <summary>Enabled while logged out.</summary>
        public static readonly ReadOnlyCollection<string> LoggedOutEnabled = new ReadOnlyCollection<string>(new[]
        {
            RibEdgeLogin, RibEdgeDebug, RibControlSheet, RibEdgeAbout,
        });

        /// <summary>Disabled when the user has no XLEdge permission (Logout stays usable).</summary>
        public static readonly ReadOnlyCollection<string> NoPermissionDisabled = new ReadOnlyCollection<string>(new[]
        {
            RibEdgeDialogBoxLauncher, RibEdgeRefresh, RibEdgeRefreshAll, RibEdgeParamRefresh, RibEdgeShowHide,
            RibEdgeIncludeOutputData, RibEdgeOptions, RibEdgeHelp, RibEdgeLogin, RibEdgeDebug, RibControlSheet, RibEdgeAbout,
            RibEdgeBroadcast,
        });
    }
}
