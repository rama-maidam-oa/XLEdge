using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using XLEdge.Helpers;
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
            EnhancedDragDropHelper.EnableWindowDrag(this);

            SelectedDate = initialDate.Date;
            CalendarControl.SelectedDate = SelectedDate;
            CalendarControl.DisplayDate = SelectedDate;
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
