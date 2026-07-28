// RETIRED - task #32 (port DPI util + WPF-UI patterns from GLSense, replace MahApps.Metro with
// WPF-UI). This file's logic was replaced by Utilities\WpfUiBootstrapper.cs and its old
// AddinModule.cs call site (MahAppsBootstrapper.Init/.PreloadResources) was updated to call
// WpfUiBootstrapper instead. This file is intentionally excluded from XLEdge.csproj's <Compile>
// list (removed there in the same change) so it no longer builds - it could not be deleted from
// disk in this environment (the mounted folder rejected the delete), so it's left here as an
// inert, clearly-marked stub instead of silently vanishing from source control history. Safe to
// delete manually once confirmed unneeded.
