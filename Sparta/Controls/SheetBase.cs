using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public class SheetBase
    {
        protected readonly ControlRoot _controlRoot;

        public SheetBase(Worksheet sheet)
        {
            _controlRoot = new ControlRoot(sheet);
        }
    }
}
