using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Sparta.Controls
{
    public class RangePainter
    {
        private Worksheet sheet;

        public readonly Value2Property Value2Property = new Value2Property();
        public readonly GenericProperty<bool> MergeCellsProperty = new GenericProperty<bool>((r, value) => r.MergeCells = value);
        public readonly GenericProperty<XlVAlign> VerticalAlignmentProperty = new GenericProperty<XlVAlign>((r, value) => r.VerticalAlignment = value);
        public readonly GenericProperty<XlHAlign> HorizontalAlignmentProperty = new GenericProperty<XlHAlign>((r, value) => r.HorizontalAlignment = value);
        public readonly GenericProperty<ExcelColor> InteriorColorProperty = new GenericProperty<ExcelColor>((r, value) => r.Interior.ColorIndex = value);

        public RangePainter(Worksheet sheet)
        {
            this.sheet = sheet;
        }

        IEnumerable<IRangeProperty> Properties
        {
            get
            {
                yield return MergeCellsProperty;
                yield return Value2Property;
                yield return HorizontalAlignmentProperty;
                yield return VerticalAlignmentProperty;
                yield return InteriorColorProperty;
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

        public XlVAlign VerticalAlignment
        {
            get { return VerticalAlignmentProperty.Value; }
            set { VerticalAlignmentProperty.Value = value; }
        }

        public XlHAlign HorizontalAlignment
        {
            get { return HorizontalAlignmentProperty.Value; }
            set { HorizontalAlignmentProperty.Value = value; }
        }

        public ExcelColor InteriorColor
        {
            get { return InteriorColorProperty.Value; }
            set { InteriorColorProperty.Value = value; }
        }
    }
}
