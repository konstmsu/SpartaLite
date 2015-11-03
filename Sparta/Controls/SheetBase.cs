using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public class SheetBase
    {
        readonly Worksheet _sheet;
        protected readonly ControlRoot ControlRoot;

        public SheetBase(Worksheet sheet)
        {
            _sheet = sheet;
            ControlRoot = new ControlRoot(sheet);
        }

        public void Run()
        {
            ControlRoot.Paint();
        }
    }
}
