using Microsoft.Office.Interop.Excel;

namespace Sparta.Controls
{
    public class ContrainerControl : IControl
    {
        public readonly ControlCollection Children = new ControlCollection();

        public virtual void Paint()
        {
            Children.Paint();
        }

        public void BeforeDoubleClick(Range target, HandledIndicator handled)
        {
            Children.OnBeforeDoubleClick(target, handled);
        }

        public Range Anchor { get; set; }

        public Range NarrowDownEventRange(Range target)
        {
            return target;
        }

        public void OnChange(Range target)
        {
            Children.OnChange(target);
        }
    }
}