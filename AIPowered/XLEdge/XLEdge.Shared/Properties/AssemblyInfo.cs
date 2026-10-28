using System.Reflection;
using System.Runtime.InteropServices;

// Version attributes come from the linked ..\XLEdgeContractsVersion.cs.
[assembly: AssemblyTitle("XLEdge.Shared")]
[assembly: AssemblyDescription("Logging, paths and release history owned by the XLEdge shell and used by the XLEdge core")]
[assembly: AssemblyCompany("Orbit Analytics")]
[assembly: AssemblyProduct("XLEdge")]
[assembly: AssemblyCopyright("Copyright © Orbit Analytics")]
[assembly: ComVisible(false)]
[assembly: Guid("0561F9D8-A62C-4EAC-8C5C-05CEF939977C")]

// Unit tests reach the lock name and the short-timeout ReadAll overload of ReleaseHistoryStore.
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("XLEdge.Tests, PublicKey=0024000004800000940000000602000000240000525341310004000001000100092eecda9a86dea143fb4b3615314e6f0fc7ce86d8d21ef6bcdec7f65f21c6e4c18395052a1e710e7d40f6e9c2e7e090e05043d26c3e72d6436706b037e58d3a1fae31f0f1c809b1fb51ff6ec604ddb7fc837ea37abf5dda333048ed2ab2c5d3d95e181feb5c95d533671b98676b0b90e8925bd15ca65916634db667fcd7d7cd")]
