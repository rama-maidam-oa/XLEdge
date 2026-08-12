using System;
using System.Windows;
using System.Windows.Input;
using XLEdge.Utilities;

namespace XLEdge.Views
{
    /// <summary>
    /// Interaction logic for XLEdgeLoginDetails.xaml
    /// </summary>
    public partial class XLEdgeLoginDetails : DpiAwareWindow
    {
        public XLEdgeLoginDetails()
        {
            InitializeComponent();
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
                LogUtility.LogException(ex, "XLEdgeLoginDetails.TitleBar_MouseLeftButtonDown");
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
