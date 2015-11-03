using System;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using Sparta.Engine.Utils;

namespace Sparta.Controls
{
    public partial class DateEditorPopupView
    {
        public DateEditorPopupView()
        {
            InitializeComponent();
        }

        public event Action ValueSelected;

        void CalendarDayButton_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            ValueSelected.Raise();
        }
    }
}
