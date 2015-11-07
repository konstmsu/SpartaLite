using System;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public class Value2Property : IRangeProperty
    {
        object _value2;

        public object Value2
        {
            get { return _value2; }
            set
            {
                if (value == null || value is int || value is double || value is string || value is DateTime || value is decimal)
                    _value2 = value;
                else
                    throw new NotSupportedException($"Value2 can't have type {value.GetType()}");
            }
        }

        public void Paint(Range range)
        {
            range.Value2 = Value2;
        }
    }
}
