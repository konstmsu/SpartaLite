using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Sparta.Utils
{
    public static class Popup
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetCursorPos(out POINT lpPoint);

        [StructLayout(LayoutKind.Sequential)]
        struct POINT
        {
            public int X;
            public int Y;
        }

        public static bool? ShowDialog(Func<Window, UserControl> getContent)
        {
            var window = new Window
            {
                SizeToContent = SizeToContent.WidthAndHeight,
                WindowStyle = WindowStyle.ToolWindow
            };
            window.Content = getContent(window);
            window.KeyUp += (sender, args) =>
            {
                if (args.Key == Key.Escape)
                    window.DialogResult = false;
            };

            window.Loaded += delegate
            {
                POINT mouse;
                if (GetCursorPos(out mouse))
                {
                    window.WindowStartupLocation = WindowStartupLocation.Manual;
                    window.Top = mouse.Y - window.Height / 2;
                    window.Left = mouse.X - window.Width / 2;
                }
            };
            
            return window.ShowDialog();
        }
    }
}
