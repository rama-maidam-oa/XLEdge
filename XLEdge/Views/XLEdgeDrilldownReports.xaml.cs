using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;
using XLEdge.Helpers;
using XLEdge.Utilities;

namespace XLEdge.Views
{
    /// <summary>
    /// Interaction logic for XLEdgeDrilldownReports.xaml
    /// </summary>
    public partial class XLEdgeDrilldownReports : DpiAwareWindow
    {
        private const double WindowMinWidth = 300;
        private const double WindowMinHeight = 230;
        private const double WindowMaxWidth = 700;
        private const double WindowMaxHeight = 700;
        private const double HorizontalPadding = 50;
        private const double VerticalPadding = 24;
        private const double HeaderHeight = 48;
        private const double FooterHeight = 58;
        private const double ItemHeight = 28;

        private readonly System.Collections.Generic.Dictionary<string, string> rptDict = new System.Collections.Generic.Dictionary<string, string>();
        private readonly ObservableCollection<DrillReportItem> reports = new ObservableCollection<DrillReportItem>();
        private System.Collections.Generic.IList<string> drillRptsList = new System.Collections.Generic.List<string>();
        private bool isLoaded;
        private bool suppressSelectionSync;

        public ObservableCollection<DrillReportItem> Reports => reports;

        public System.Collections.Generic.IList<string> DrillRptsList
        {
            get => drillRptsList;
            set
            {
                drillRptsList = value ?? new System.Collections.Generic.List<string>();
                if (isLoaded)
                {
                    RefreshReports();
                    AdjustWindowSize();
                }
            }
        }

        public string DrillSelRpt { get; private set; } = string.Empty;

        public XLEdgeDrilldownReports()
        {
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);
            DataContext = this;
            Loaded += XLEdgeDrilldownReports_Loaded;
        }

        public XLEdgeDrilldownReports(System.Collections.Generic.IEnumerable<string> drillReports) : this()
        {
            DrillRptsList = drillReports?.ToList() ?? new System.Collections.Generic.List<string>();
        }

        private void XLEdgeDrilldownReports_Loaded(object sender, RoutedEventArgs e)
        {
            if (isLoaded)
            {
                return;
            }

            isLoaded = true;
            RefreshReports();
            AdjustWindowSize();
            PositionWindowNearActiveCell();
        }

        private void RefreshReports()
        {
            rptDict.Clear();
            reports.Clear();

            if (DrillRptsList == null || DrillRptsList.Count == 0)
            {
                return;
            }

            foreach (string item in DrillRptsList)
            {
                if (string.IsNullOrWhiteSpace(item))
                {
                    continue;
                }

                string[] parts = item.Split('|');
                if (parts.Length < 3)
                {
                    continue;
                }

                string name = parts[2];
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (!rptDict.ContainsKey(name))
                {
                    rptDict.Add(name, item);
                }

                reports.Add(new DrillReportItem(name, item, reports.Count + 1, OnReportChecked));
            }
        }

        private void AdjustWindowSize()
        {
            double maxItemWidth = WindowMinWidth - HorizontalPadding;

            foreach (DrillReportItem item in reports)
            {
                maxItemWidth = System.Math.Max(maxItemWidth, MeasureTextWidth(item.DisplayName));
            }

            double listHeight = System.Math.Max(ItemHeight, reports.Count * ItemHeight + 8);
            double maxListHeight = System.Math.Max(WindowMinHeight, WindowMaxHeight - HeaderHeight - FooterHeight - VerticalPadding);
            listHeight = System.Math.Min(listHeight, maxListHeight);

            Width = System.Math.Min(WindowMaxWidth, System.Math.Max(WindowMinWidth, maxItemWidth + HorizontalPadding + 40));
            Height = System.Math.Min(WindowMaxHeight, System.Math.Max(WindowMinHeight, HeaderHeight + FooterHeight + VerticalPadding + listHeight));

        }

        private double MeasureTextWidth(string text)
        {
            FontFamily fontFamily = new FontFamily("Segoe UI");
            FontStyle fontStyle = FontStyles.Normal;
            FontWeight fontWeight = FontWeights.Normal;
            FontStretch fontStretch = FontStretches.Normal;
            double fontSize = 13;

            Typeface typeface = new Typeface(fontFamily, fontStyle, fontWeight, fontStretch);
            double pixelsPerDip = VisualTreeHelper.GetDpi(this).PixelsPerDip;
            FormattedText formattedText = new FormattedText(
                text ?? string.Empty,
                CultureInfo.CurrentUICulture,
                FlowDirection.LeftToRight,
                typeface,
                fontSize,
                Brushes.Black,
                pixelsPerDip);

            return formattedText.WidthIncludingTrailingWhitespace;
        }

        private void PositionWindowNearActiveCell()
        {
            try
            {
                if (!ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp))
                {
                    WindowStartupLocation = WindowStartupLocation.CenterScreen;
                    return;
                }

                SetExcelOwner(new System.IntPtr(excelApp.Hwnd));

                Excel.Range cell = excelApp.ActiveCell;
                Excel.Window activeWindow = excelApp.ActiveWindow;
                if (cell == null || activeWindow == null)
                {
                    WindowStartupLocation = WindowStartupLocation.CenterScreen;
                    return;
                }

                PresentationSource source = PresentationSource.FromVisual(this);
                double scaleX = source?.CompositionTarget?.TransformFromDevice.M11 ?? 1.0;
                double scaleY = source?.CompositionTarget?.TransformFromDevice.M22 ?? 1.0;

                double cellLeftPoints = System.Convert.ToDouble(cell.Left);
                double cellTopPoints = System.Convert.ToDouble(cell.Top);
                double cellHeightPoints = System.Convert.ToDouble(cell.Height);

                double cellLeft = activeWindow.PointsToScreenPixelsX((int)System.Math.Round(cellLeftPoints)) / scaleX;
                double cellTop = activeWindow.PointsToScreenPixelsY((int)System.Math.Round(cellTopPoints + cellHeightPoints)) / scaleY;

                Left = cellLeft;
                Top = cellTop;
                WindowStartupLocation = WindowStartupLocation.Manual;

                ClampToWorkArea();
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Drilldown report window positioning failed: {ex.Message}");
                WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
        }

        private void ClampToWorkArea()
        {
            Rect workArea = SystemParameters.WorkArea;

            if (Left + Width > workArea.Right)
            {
                Left = workArea.Right - Width;
            }

            if (Top + Height > workArea.Bottom)
            {
                Top = workArea.Bottom - Height;
            }

                Left = System.Math.Max(Left, workArea.Left);
                Top = System.Math.Max(Top, workArea.Top);
        }

        private void OnReportChecked(DrillReportItem selectedItem)
        {
            if (suppressSelectionSync)
            {
                return;
            }

            try
            {
                suppressSelectionSync = true;

                foreach (DrillReportItem item in reports)
                {
                    if (!ReferenceEquals(item, selectedItem))
                    {
                        item.IsChecked = false;
                    }
                }
            }
            finally
            {
                suppressSelectionSync = false;
            }
        }

        private void CmdClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void CmdExecute_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DrillReportItem selectedItem = reports.FirstOrDefault(item => item.IsChecked);

                if (selectedItem == null)
                {
                    DrillSelRpt = string.Empty;
                }
                else
                {
                    string reportValue = rptDict.TryGetValue(selectedItem.DisplayName, out string mappedValue)
                        ? mappedValue
                        : selectedItem.OriginalValue;

                    DrillSelRpt = $"{reportValue}|{selectedItem.Index}";
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogError(ex.ToString());
            }
            finally
            {
                Close();
            }
        }

        private void HeaderBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        public sealed class DrillReportItem : INotifyPropertyChanged
        {
            private readonly Action<DrillReportItem> onChecked;
            private bool isChecked;

            public DrillReportItem(string displayName, string originalValue, int index, Action<DrillReportItem> onChecked)
            {
                DisplayName = displayName;
                OriginalValue = originalValue;
                Index = index;
                this.onChecked = onChecked;
            }

            public string DisplayName { get; }

            public string OriginalValue { get; }

            public int Index { get; }

            public bool IsChecked
            {
                get => isChecked;
                set
                {
                    if (isChecked == value)
                    {
                        return;
                    }

                    isChecked = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsChecked)));

                    if (value)
                    {
                        onChecked?.Invoke(this);
                    }
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
        }
    }
}
