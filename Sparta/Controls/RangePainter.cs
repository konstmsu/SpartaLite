using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Sparta.Controls
{
    public class RangePainter
    {
        public readonly Value2Property Value2Property = new Value2Property();
        public readonly GenericProperty<string> NumberFormatProperty = new GenericProperty<string>((r, value) => r.NumberFormat = value);
        public readonly GenericProperty<bool> MergeCellsProperty = new GenericProperty<bool>((r, value) => r.MergeCells = value);
        public readonly GenericProperty<XlVAlign?> VerticalAlignmentProperty = new GenericProperty<XlVAlign?>((r, value) =>
        {
            if (value != null)
                r.VerticalAlignment = value;
        });
        public readonly GenericProperty<XlHAlign?> HorizontalAlignmentProperty = new GenericProperty<XlHAlign?>((r, value) =>
        {
            if (value != null)
                r.HorizontalAlignment = value;
        });
        public readonly GenericProperty<ExcelColor> InteriorColorProperty = new GenericProperty<ExcelColor>((r, value) => value?.Apply(r.Interior));
        public readonly GenericProperty<ExcelColor> FontColorProperty = new GenericProperty<ExcelColor>((r, value) => value?.Apply(r.Font));
        public readonly GenericProperty<bool> IsBoldProperty = new GenericProperty<bool>((r, value) => r.Font.Bold = value);
        public readonly BorderProperty Border = new BorderProperty();
        public readonly ValidationProperty Validation = new ValidationProperty();

        IEnumerable<IRangeProperty> Properties
        {
            get
            {
                yield return MergeCellsProperty;
                yield return Value2Property;
                yield return NumberFormatProperty;
                yield return HorizontalAlignmentProperty;
                yield return VerticalAlignmentProperty;
                yield return InteriorColorProperty;
                yield return FontColorProperty;
                yield return IsBoldProperty;
                yield return Border;
                yield return Validation;
            }
        }

        public void Paint(Range range)
        {
            foreach (var p in Properties)
                p.Paint(range);
        }

        public object Value2
        {
            get { return Value2Property.Value2; }
            set { Value2Property.Value2 = value; }
        }

        public bool MergeCells
        {
            get { return MergeCellsProperty.Value; }
            set { MergeCellsProperty.Value = value; }
        }

        public XlVAlign? VerticalAlignment
        {
            get { return VerticalAlignmentProperty.Value; }
            set { VerticalAlignmentProperty.Value = value; }
        }

        public XlHAlign? HorizontalAlignment
        {
            get { return HorizontalAlignmentProperty.Value; }
            set { HorizontalAlignmentProperty.Value = value; }
        }

        public ExcelColor InteriorColor
        {
            get { return InteriorColorProperty.Value; }
            set { InteriorColorProperty.Value = value; }
        }

        public ExcelColor FontColor
        {
            get { return FontColorProperty.Value; }
            set { FontColorProperty.Value = value; }
        }

        public bool IsBold
        {
            get { return IsBoldProperty.Value; }
            set { IsBoldProperty.Value = value; }
        }

        public string NumberFormat
        {
            get { return NumberFormatProperty.Value; }
            set { NumberFormatProperty.Value = value; }
        }
    }
}
