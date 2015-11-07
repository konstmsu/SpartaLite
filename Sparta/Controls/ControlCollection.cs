using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Sparta.Engine.Utils;

namespace Sparta.Controls
{
    public static class ControlCollectionExtensions
    {
        public static void OnBeforeDoubleClick(this IEnumerable<IControl> controls, Range target, HandledIndicator handled)
        {
            foreach (var control in controls)
            {
                var range = control.NarrowDownEventRange(target);

                if (range != null)
                    control.BeforeDoubleClick(range, handled);
            }
        }

        public static void OnChange(this IEnumerable<IControl> controls, Range target)
        {
            controls.ForEachAggregatingExceptions(control =>
            {
                var range = control.NarrowDownEventRange(target);

                if (range != null)
                    control.OnChange(range);
            });
        }

        public static void Paint(this IEnumerable<IControl> controls)
        {
            foreach (var control in controls)
                control.Paint();
        }
    }
}
