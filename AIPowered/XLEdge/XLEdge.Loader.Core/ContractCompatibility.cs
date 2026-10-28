using XLEdge.Contracts;

namespace XLEdge.Loader.Core
{
    /// <summary>
    /// Whether this shell can load a core release built against a given contract version: from 1 up to the shell's own
    /// ContractInfo.Version (older cores are fine — Release History rollback). 0 means the release did not record one.
    /// </summary>
    public static class ContractCompatibility
    {
        public static bool IsLoadable(int contractVersion) => contractVersion >= 1 && contractVersion <= ContractInfo.Version;

        /// <summary>Why a release cannot be loaded, or null when it can.</summary>
        public static string RefusalReason(int contractVersion)
        {
            if (IsLoadable(contractVersion)) return null;
            if (contractVersion > ContractInfo.Version)
                return $"needs a newer XLEdge installer (contract {contractVersion}, this installer supports up to {ContractInfo.Version})";
            return "built without a contract version";
        }
    }
}
