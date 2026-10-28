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
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("XLEdge.Tests, PublicKey=00240000048000009400000006020000002400005253413100040000010001002979e5ced75c98f450e481b29180bdc351de1597470c68a7a6bd4dfed2537747f541314619e20b800c397cfd95c5ad35b36734d32e5d9c4a5f3090e8f69a2451e9ad356e0730469f59fe2efa324fe23deea3d0175a3cd330da3880daef048a2ad27b3877345155ebc70ee3513224817dc9a57f2e259b379116f8c87e3ab42db5")]
