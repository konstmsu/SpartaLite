namespace Sparta.Controls
{
    public class ExcelColor
    {
        public static readonly ExcelColor DarkGray = new ExcelColor(123);
        readonly int _color;

        public ExcelColor(int color)
        {
            _color = color;
        }

        public void Apply()
        {

        }
    }
}
