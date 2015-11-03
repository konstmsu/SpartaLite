using System;
using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public class SheetFactory
    {
        readonly Application _application;

        public SheetFactory(Application application)
        {
            _application = application;
        }

        public void ShowSheet<T>(Func<Worksheet, T> getHandler)
            where T: SheetBase
        {
            var sheet = (Worksheet)_application.Worksheets.Add();
            var handler = getHandler(sheet);
            handler.Run();
        }
    }
}
