using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public interface IControl
    {
        void Paint();
        void BeforeDoubleClick(Range target, ref bool handled);
    }
}
