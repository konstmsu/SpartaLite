using System;
using System.Windows;

namespace Sparta.Controls
{
    public partial class DateEditorPopupView
    {
        public DateEditorPopupView()
        {
            InitializeComponent();
        }

        public event Action ValueSelected;

        void CalendarDayButton_DoubleClick(object sender, RoutedEventArgs e)
        {
            ValueSelected?.Invoke();
        }
    }
}
