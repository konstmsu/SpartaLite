using Microsoft.Office.Interop.Excel;
using System;

namespace Sparta.Controls
{
    public class ExcelColor
    {
        public static readonly ExcelColor Red = FromRgb(255, 0, 0);
        public static readonly ExcelColor Green = FromRgb(0, 255, 0);
        public static readonly ExcelColor Blue = FromRgb(0, 0, 255);
        public static readonly ExcelColor LightBlue = FromRgb(60, 60, 255);
        public static readonly ExcelColor DarkGray = FromRgb(40, 40, 40);
        public static readonly ExcelColor LightGray = FromRgb(160, 160, 160);

        public readonly int Code;

        public static ExcelColor FromRgb(int red, int green, int blue)
        {
            if (red < 0 || red > 255) throw new ArgumentOutOfRangeException();
            if (green < 0 || green > 255) throw new ArgumentOutOfRangeException();
            if (blue < 0 || blue > 255) throw new ArgumentOutOfRangeException();

            return new ExcelColor(red + green * 256 + blue * 256 * 256);
        }

        ExcelColor(int code)
        {
            Code = code;
        }

        public void Apply(Interior interior)
        {
            interior.Color = Code;
        }

        public void Apply(Font font)
        {
            font.Color = Code;
        }
    }
}
