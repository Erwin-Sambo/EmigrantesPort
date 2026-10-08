using Microsoft.Maui.Controls.Shapes;
namespace EmigrantesPort;

internal static class Design
{
    internal static readonly Color Ink = Color.FromArgb("#253B35"), Muted = Color.FromArgb("#71837B"), Green = Color.FromArgb("#173F35"), Line = Color.FromArgb("#E2E8E1"), Canvas = Color.FromArgb("#F7F8F4");
    internal static Label Label(string text, double size = 14, Color? color = null, bool bold = false) => new()
    { Text = text, FontSize = size, TextColor = color ?? Ink, FontFamily = bold ? "OpenSansSemibold" : "OpenSansRegular", VerticalOptions = LayoutOptions.Center };
    internal static Button Button(string text, bool primary = false) => new()
    { Text = text, BackgroundColor = primary ? Green : Colors.White, TextColor = primary ? Colors.White : Green, BorderColor = primary ? Green : Line, BorderWidth = 1, Padding = new Thickness(17, 11), HeightRequest = 44 };
    internal static Border Card(View child, Thickness? padding = null, Color? background = null) => new()
    { Content = child, Padding = padding ?? new Thickness(22), BackgroundColor = background ?? Colors.White, Stroke = Line, StrokeThickness = 1, StrokeShape = new RoundRectangle { CornerRadius = 12 } };
    internal static Grid Grid(string columns, string rows = "Auto")
    {
        static GridLength Length(string value) => value == "Auto" ? GridLength.Auto : value.EndsWith('*') ? new GridLength(value == "*" ? 1 : double.Parse(value[..^1], System.Globalization.CultureInfo.InvariantCulture), GridUnitType.Star) : new GridLength(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture));
        var grid = new Grid();
        foreach (var c in columns.Split(',')) grid.ColumnDefinitions.Add(new ColumnDefinition(Length(c)));
        foreach (var r in rows.Split(',')) grid.RowDefinitions.Add(new RowDefinition(Length(r)));
        return grid;
    }
    internal static View Field(string label, View input, string? hint = null)
    {
        var layout = new VerticalStackLayout { Spacing = 7 };
        layout.Add(Label(label, 12, Muted, true)); layout.Add(Card(input, new Thickness(12, 3)));
        if (hint is not null) layout.Add(Label(hint, 11, Muted)); return layout;
    }
}
