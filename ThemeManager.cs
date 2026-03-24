using System.Drawing;

namespace OfficeInstaller
{

    public enum Theme
    {
        Light,
        Dark
    }

    public static class ThemeManager
    {

        public static Theme CurrentTheme = Theme.Dark;

        public static Color BackColor =>
            CurrentTheme == Theme.Dark
            ? Color.FromArgb(24, 24, 24)
            : Color.White;

        public static Color ForeColor =>
            CurrentTheme == Theme.Dark
            ? Color.White
            : Color.Black;

        public static Color TitleBar =>
            CurrentTheme == Theme.Dark
            ? Color.FromArgb(32, 32, 32)
            : Color.FromArgb(240, 240, 240);

    }

}