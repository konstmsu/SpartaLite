using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Sparta.Utils
{
    public static class ApplicationExtensions
    {
        public static IDisposable SetTemporarily(this Application application, bool? screenUpdating = null, bool? enableEvents = null)
        {
            var changes = new List<IDisposable>();

            if (screenUpdating.HasValue)
                changes.Add(SetTemporarily(screenUpdating.Value, () => application.ScreenUpdating, value => application.ScreenUpdating = value));

            if (enableEvents.HasValue)
                changes.Add(SetTemporarily(enableEvents.Value, () => application.EnableEvents, value => application.EnableEvents = value));

            return new CompositeDisposable(changes);
        }

        static IDisposable SetTemporarily(bool value, Func<bool> getter, Action<bool> setter)
        {
            var oldValue = getter();
            setter(value);
            return new Disposable(() => setter(oldValue));
        }
    }
}
