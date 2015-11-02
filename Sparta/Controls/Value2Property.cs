using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public class Value2Property : IRangeProperty
    {
        public object Value2 { get; set; }

        public void Paint(Range range)
        {
            range.Value2 = Value2;
        }
    }
}
