using System;
using System.Windows;
using System.Windows.Input;
using XLEdge.Utilities;

namespace XLEdge.Views
{
    /// <summary>
    /// Interaction logic for XLEdgeCalendar.xaml
    /// </summary>
    public partial class XLEdgeCalendar : DpiAwareWindow    
    {
        public DateTime SelectedDate { get; private set; }
        public XLEdgeCalendar(DateTime initialDate)
        {
            InitializeComponent();

            SelectedDate = initialDate.Date;
            CalendarControl.SelectedDate = SelectedDate;
            CalendarControl.DisplayDate = SelectedDate;
        }

        // Replaces EnhancedDragDropHelper.EnableWindowDrag(this) now that the window has a real
        // title bar (WindowStyle="SingleBorderWindow" + ExtendsContentIntoTitleBar="True").
        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    this.DragMove();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "XLEdgeCalendar.TitleBar_MouseLeftButtonDown");
            }
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            SelectedDate = CalendarControl.SelectedDate ?? DateTime.Today;
            DialogResult = true;
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
