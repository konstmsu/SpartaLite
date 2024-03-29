﻿using System;
using Microsoft.Office.Interop.Excel;
using Sparta.Utils;

namespace Sparta.Controls
{
    public class DateEditorControl : IControl
    {
        public DateTime Value
        {
            get { return (DateTime)_painter.Value2; }
            set { _painter.Value2 = value; }
        }
        readonly RangePainter _painter;

        public DateEditorControl()
        {
            _painter = new RangePainter
            {
                NumberFormat = KnownFormats.Date
            };
        }

        public void Paint()
        {
            _painter.InteriorColor = IsDisabled ? SpartaColors.DisabledControlInterior : SpartaColors.DefaultControlInterior;
            _painter.FontColor = IsDisabled ? SpartaColors.DisabledControlFont : SpartaColors.ValueWithPopup;

            _painter.Paint(Anchor);
        }

        public void BeforeDoubleClick(Range target, HandledIndicator handled)
        {
            handled.MarkHandled();

            if (IsDisabled)
                return;

            var viewModel = new DateEditorPopupViewModel { Value = Value };

            var dialogResult = Popup.ShowDialog(w =>
            {
                var view = new DateEditorPopupView { DataContext = viewModel };
                view.ValueSelected += () => w.DialogResult = true;
                return view;
            });

            if (dialogResult == true)
                Value = viewModel.Value;
        }

        public Range Anchor { get; set; }
        public bool IsDisabled { get; set; }

        public Range NarrowDownEventRange(Range target)
        {
            return target.GetIntersection(Anchor);
        }

        public void OnChange(Range target)
        {
            if (IsDisabled)
                return;

            Value = CoerceToDateTime(target.Value2);
        }

        static DateTime CoerceToDateTime(object value)
        {
            return CoerceTo<DateTime>.Value(value)
                .Type<DateTime>(v => v)
                .Type<double>(v => DateTime.FromOADate(v))
                .Type<int>(v => DateTime.FromOADate(v))
                .ElseThrow();
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