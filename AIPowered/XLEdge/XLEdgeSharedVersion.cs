// XLEdgeSharedVersion.cs - solution root
//
// Product version of XLEdge. Linked (not copied) into XLEdge (shell), XLEdge.Addin.Core
// and XLEdge.Build through <Compile Include="..\XLEdgeSharedVersion.cs"> with
// <Link>Properties\XLEdgeSharedVersion.cs</Link>, so a release bumps the version here only.
//
// XLEdge.Contracts, XLEdge.Shared and XLEdge.Loader.Core do NOT use this file: their
// AssemblyVersion is frozen in XLEdgeContractsVersion.cs so that a core-only release
// never changes the identity of the types the shell and the core exchange.
//
// Do not add AssemblyVersion/AssemblyFileVersion to any project's Properties\AssemblyInfo.cs.
using System.Reflection;

[assembly: AssemblyVersion("11.1.2.0")]
[assembly: AssemblyFileVersion("11.1.2.0")]
