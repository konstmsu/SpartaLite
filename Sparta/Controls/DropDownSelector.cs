using System;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Sparta.Engine.Utils;
using Sparta.Utils;

namespace Sparta.Controls
{
    public class DropDownSelector : IControl
    {
        string _selectedValue
        {
            get { return (string)_painter.Value2; }
            set { _painter.Value2 = value; }
        }

        public string SelectedValue
        {
            get { return _selectedValue; }
            set
            {
                if (_selectedValue == value)
                    return;

                _selectedValue = value;
                SelectedValueChanged?.Invoke();
            }
        }

        public event System.Action SelectedValueChanged;

        public ReadOnlyCollection<string> Values
        {
            get { return _values; }
            set { _values = value; }
        }

        ReadOnlyCollection<string> _values
        {
            get { return _painter.Validation.List; }
            set { _painter.Validation.List = value; }
        }

        public Range Anchor { get; set; }

        readonly RangePainter _painter;

        public DropDownSelector()
        {
            _painter = new RangePainter();
        }

        public void Paint()
        {
            _painter.Paint(Anchor);
        }

        public void BeforeDoubleClick(Range target, HandledIndicator handled)
        {
            var selectedIndex = Values.IndexOf(SelectedValue);
            SelectedValue = Values[(selectedIndex + 1) % Values.Count];
            handled.MarkHandled();
        }

        public Range NarrowDownEventRange(Range target)
        {
            return target.GetIntersection(Anchor);
        }

        public void OnChange(Range target)
        {
            var newValue = (string)target.Value2;

            var bestMatch = Values.Where(v => v.Equals(newValue, StringComparison.OrdinalIgnoreCase)).ToList();

            if (bestMatch.Any())
            {
                SelectedValue = bestMatch.First();
                return;
            }

            bestMatch = Values.Where(v => v.StartsWith(newValue, StringComparison.OrdinalIgnoreCase)).ToList();

            if (bestMatch.Any())
            {
                SelectedValue = bestMatch.First();
                return;
            }

            throw new ApplicationException($"Could not match input '{newValue}' with any of the DropDown values [{Values.JoinStrings(", ")}]");
        }
    }
}
