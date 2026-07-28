using System.Windows;
using XLEdge.Helpers;
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
            EnhancedDragDropHelper.EnableWindowDrag(this);
        }
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
