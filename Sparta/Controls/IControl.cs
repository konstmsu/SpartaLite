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
        Range Anchor { get; set; }
        void Paint();
        void BeforeDoubleClick(Range target, HandledIndicator handled);
        Range NarrowDownEventRange(Range target);
        void OnChange(Range target);
    }
}
