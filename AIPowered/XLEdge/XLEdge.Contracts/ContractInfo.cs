namespace XLEdge.Contracts
{
    /// <summary>
    /// Version of this contract (the interfaces and data types in XLEdge.Contracts).
    /// <para>
    /// Increment it on every contract change; every contract change ships with a new MSI. Each core
    /// release records the contract version it was built against (contractVersion in its manifest.json),
    /// and the loader only loads a release whose contract version is not newer than the shell's.
    /// Older cores are allowed (Release History rollback): the shell calls members added after their
    /// version inside try/catch.
    /// </para>
    /// </summary>
    public static class ContractInfo
    {
        public const int Version = 1;
    }
}
