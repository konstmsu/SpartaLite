using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public class SheetBase
    {
        protected readonly ControlRoot ControlRoot;

        public SheetBase(Worksheet sheet)
        {
            ControlRoot = new ControlRoot(sheet);
        }

        public void Run()
        {
            ControlRoot.Paint();
        }
    }
}
