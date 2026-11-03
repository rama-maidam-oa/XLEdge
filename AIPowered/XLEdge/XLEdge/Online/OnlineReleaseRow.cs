using System.ComponentModel;

namespace XLEdge.Online
{
    /// <summary>One server release in the reload picker's Online list (GLSense's OnlineReleaseRow, XLEdge-named).</summary>
    internal sealed class OnlineReleaseRow : INotifyPropertyChanged
    {
        private bool _isChecked;

        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public bool Mandatory { get; set; }
        public string FolderName { get; set; }
        public string FileName { get; set; }
        public string Checksum { get; set; }
        public string Notes { get; set; }

        public bool IsCurrentlyLoaded { get; set; }
        public bool IsAlreadyDownloaded { get; set; }

        /// <summary>Neither loaded nor already downloaded: the only rows that can be ticked.</summary>
        public bool IsSelectable => !IsCurrentlyLoaded && !IsAlreadyDownloaded;

        public string StatusText => IsCurrentlyLoaded ? "Currently loaded" : IsAlreadyDownloaded ? "Already downloaded" : "New";

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
