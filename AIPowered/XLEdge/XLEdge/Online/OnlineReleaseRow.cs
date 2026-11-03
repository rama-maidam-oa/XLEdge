using System.ComponentModel;
using XLEdge.Shared;

namespace XLEdge.Online
{
    /// <summary>One server release in the reload picker's Online list (GLSense's OnlineReleaseRow, XLEdge-named).</summary>
    internal sealed class OnlineReleaseRow : INotifyPropertyChanged
    {
        private bool _isChecked;

        public string Version { get; set; }
        /// <summary>The version as shown (11.1.3, not 11.1.3.0).</summary>
        public string DisplayVersion => VersionText.Display(Version);
        public string ReleaseDate { get; set; }
        public bool Mandatory { get; set; }
        public string FolderName { get; set; }
        public string FileName { get; set; }
        public string Checksum { get; set; }
        public string Notes { get; set; }

        public bool IsCurrentlyLoaded { get; set; }
        public bool IsAlreadyDownloaded { get; set; }

        /// <summary>Built for a newer contract than this shell's (it needs a newer XLEdge install).</summary>
        public bool NeedsNewerInstall { get; set; }

        /// <summary>Neither loaded, already downloaded nor too new for this shell: the only rows that can be ticked.</summary>
        public bool IsSelectable => !IsCurrentlyLoaded && !IsAlreadyDownloaded && !NeedsNewerInstall;

        public string StatusText =>
            IsCurrentlyLoaded ? "Currently loaded" : IsAlreadyDownloaded ? "Already downloaded" : NeedsNewerInstall ? "Needs a newer XLEdge install" : "New";

        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                bool checkedNow = value && IsSelectable;
                if (_isChecked == checkedNow) return;
                _isChecked = checkedNow;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsChecked)));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
