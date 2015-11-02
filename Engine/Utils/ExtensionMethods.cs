using System;

namespace Sparta.Engine.Utils
{
    public static class ExtensionMethods
    {
        public static void Raise(this Action action)
        {
            action?.Invoke();
        }
    }
}
