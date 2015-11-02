using System;
using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public class GenericProperty<T> : IRangeProperty
    {
        public T Value;
        readonly Action<Range, T> _paint;
            
        public GenericProperty(Action<Range, T> paint)
        {
            _paint = paint;
        }

        public void Paint(Range range)
        {
            _paint(range, Value);
        }
    }
}
