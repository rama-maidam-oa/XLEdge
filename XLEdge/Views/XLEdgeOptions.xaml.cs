using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using XLEdge.Helpers;
using XLEdge.Utilities;

namespace XLEdge.Views
{
    public partial class XLEdgeOptions : DpiAwareWindow, INotifyPropertyChanged
    {
        private bool parameterValuesInSameSheet;
        private bool downloadScheduledOutputsToExistingSheets;
        private bool syncWithReportDefinition;
        private bool overrideSheetNameForScheduledOutputs;
        private bool showCalendarControl;
        private bool showSegmentSelectionWindow;
        private bool overrideFormats;

        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler PreferencesApplied;

        public bool ParameterValuesInSameSheet
        {
            get => parameterValuesInSameSheet;
            set
            {
                if (parameterValuesInSameSheet != value)
                {
                    parameterValuesInSameSheet = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool DownloadScheduledOutputsToExistingSheets
        {
            get => downloadScheduledOutputsToExistingSheets;
            set
            {
                if (downloadScheduledOutputsToExistingSheets != value)
                {
                    downloadScheduledOutputsToExistingSheets = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool SyncWithReportDefinition
        {
            get => syncWithReportDefinition;
            set
            {
                if (syncWithReportDefinition != value)
                {
                    syncWithReportDefinition = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool OverrideSheetNameForScheduledOutputs
        {
            get => overrideSheetNameForScheduledOutputs;
            set
            {
                if (overrideSheetNameForScheduledOutputs != value)
                {
                    overrideSheetNameForScheduledOutputs = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool ShowCalendarControl
        {
            get => showCalendarControl;
            set
            {
                if (showCalendarControl != value)
                {
                    showCalendarControl = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool ShowSegmentSelectionWindow
        {
            get => showSegmentSelectionWindow;
            set
            {
                if (showSegmentSelectionWindow != value)
                {
                    showSegmentSelectionWindow = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool OverrideFormats
        {
            get => overrideFormats;
            set
            {
                if (overrideFormats != value)
                {
                    overrideFormats = value;
                    OnPropertyChanged();
                }
            }
        }

        public XLEdgeOptions()
        {
            InitializeComponent();
            DataContext = this;

            EnhancedDragDropHelper.EnableWindowDrag(this);
            LoadFromAppState();
        }

        private void LoadFromAppState()
        {
            var appState = XLEdgeAppState.Instance;
            ParameterValuesInSameSheet = appState.ParamDataSameSheet;
            DownloadScheduledOutputsToExistingSheets = appState.SchOutputsToSameSheet;
            SyncWithReportDefinition = appState.RefreshSync;
            OverrideSheetNameForScheduledOutputs = appState.AllowSheetNameChanges;
            ShowCalendarControl = appState.ShowCalendarControl;
            ShowSegmentSelectionWindow = appState.ShowSegmentSelectionWindow;
            OverrideFormats = appState.OverrideFormats;
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            ApplyPreferencesToAppStateOnly();
            PreferencesApplied?.Invoke(this, EventArgs.Empty);
            Close();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var preferences = BuildPreferencesFromWindow();
            XLEdgePreferencesManager.Instance.Save(preferences);
            PreferencesApplied?.Invoke(this, EventArgs.Empty);
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ApplyPreferencesToAppStateOnly()
        {
            var preferences = BuildPreferencesFromWindow();
            XLEdgePreferencesManager.Instance.ApplyRuntime(preferences);
        }

        private Models.XLEdgeUserPreferences BuildPreferencesFromWindow()
        {
            return new Models.XLEdgeUserPreferences
            {
                ParameterValues = ParameterValuesInSameSheet,
                ScheduledOutputs = DownloadScheduledOutputsToExistingSheets,
                RefreshSync = SyncWithReportDefinition,
                ChangeSheetName = OverrideSheetNameForScheduledOutputs,
                CalendarCtrlDisplay = ShowCalendarControl,
                SegmentSelectionWindowDisplay = ShowSegmentSelectionWindow,
                OverrideFormats = OverrideFormats
            };
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}