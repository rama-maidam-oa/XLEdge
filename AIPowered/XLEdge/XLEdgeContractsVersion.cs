// XLEdgeContractsVersion.cs - solution root
//
// Version of the assemblies shared between the shell and the hot-swappable core:
// XLEdge.Contracts, XLEdge.Shared and XLEdge.Loader.Core (linked into each of them).
//
// AssemblyVersion is FROZEN. These assemblies are strong-named and installed next to
// XLEdge.dll by the MSI; XLEdge.Contracts and XLEdge.Shared also ship inside every core zip
// (the core's AppDomain resolves them from its release folder), while XLEdge.Loader.Core is
// always loaded from the shell's copy. If a core release changed their AssemblyVersion, the
// types created in the core's AppDomain would no longer match the shell's copies and the
// cross-domain casts would fail. Change it only together with a new MSI, and only for a
// breaking contract change.
//
// AssemblyFileVersion follows the product version so the DLLs show which release built them.
using System.Reflection;

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("11.1.3.0")]
