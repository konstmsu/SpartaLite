using System;
using Microsoft.Office.Interop.Excel;
using Sparta.Utils;

namespace Sparta.Controls
{
    public class DateEditorControl : IControl
    {
        public DateTime Value;
        readonly RangePainter _painter;

        public DateEditorControl()
        {
            _painter = new RangePainter
            {
                FontColor = SpartaColors.ValueWithPopup,
                NumberFormat = KnownFormats.Date
            };
        }

        public void Paint()
        {
            _painter.Value2 = Value;
            _painter.Paint(Anchor);
        }

        public void BeforeDoubleClick(Range target, HandledIndicator handled)
        {
            DateEditorPopupViewModel viewModel = null;

            var dialogResult = Popup.ShowDialog(w =>
            {
                var view = new DateEditorPopupView
                {
                    DataContext = viewModel = new DateEditorPopupViewModel
                    {
                        Value = Value
                    }
                };

                view.ValueSelected += () => w.DialogResult = true;

                return view;
            });

            if (dialogResult == true)
                Value = viewModel.Value;

            handled.MarkHandled();
        }

        public Range Anchor { get; set; }
        public Range NarrowDownEventRange(Range target)
        {
            return target.GetIntersection(Anchor);
        }
    }

    public class HandledIndicator
    {
        public bool IsHandled { get; private set; }
        public void MarkHandled()
        {
            IsHandled = true;
        }
    }
}