using System;
using System.Windows;
using System.Windows.Controls;

namespace Sparta.Utils
{
    public static class Popup
    {
        public static bool? ShowDialog(Func<Window, UserControl> getContent)
        {
            var window = new Window
            {
                SizeToContent = SizeToContent.WidthAndHeight
            };

            var content = getContent(window);

            window.Content = content;
            
            return window.ShowDialog();
        }
    }
}
