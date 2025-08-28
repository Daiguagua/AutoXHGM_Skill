using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrayNotify;

public class DebugOverlayWindow : Window
{
    private readonly Canvas _canvas;
    private readonly List<UIElement> _debugElements = new List<UIElement>();

    public DebugOverlayWindow()
    {
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Background = Brushes.Transparent;
        Topmost = true;
        ShowInTaskbar = false;
        Width = SystemParameters.PrimaryScreenWidth;
        Height = SystemParameters.PrimaryScreenHeight;

        _canvas = new Canvas();
        Content = _canvas;
    }

    public void AddDebugRectangle(System.Drawing.Rectangle region, System.Windows.Media.Color color, string label = "")
    {
        var rect = new Rectangle
        {
            Width = region.Width,
            Height = region.Height,
            Stroke = new SolidColorBrush(color),
            StrokeThickness = 2,
            Fill = Brushes.Transparent
        };

        Canvas.SetLeft(rect, region.X);
        Canvas.SetTop(rect, region.Y);
        _canvas.Children.Add(rect);
        _debugElements.Add(rect);

        if (!string.IsNullOrEmpty(label))
        {
            var textBlock = new TextBlock
            {
                Text = label,
                Foreground = Brushes.White,
                Background = new SolidColorBrush(Color.FromArgb(128, 0, 0, 0)),
                Padding = new Thickness(2)
            };

            Canvas.SetLeft(textBlock, region.X);
            Canvas.SetTop(textBlock, region.Y - 20);
            _canvas.Children.Add(textBlock);
            _debugElements.Add(textBlock);
        }
    }

    public void AddDebugPoint(System.Drawing.Point point, System.Windows.Media.Color color, string label = "")
    {
        var ellipse = new Ellipse
        {
            Width = 6,
            Height = 6,
            Fill = new SolidColorBrush(color),
            Stroke = Brushes.White,
            StrokeThickness = 1
        };

        Canvas.SetLeft(ellipse, point.X - 3);
        Canvas.SetTop(ellipse, point.Y - 3);
        _canvas.Children.Add(ellipse);
        _debugElements.Add(ellipse);

        if (!string.IsNullOrEmpty(label))
        {
            var textBlock = new TextBlock
            {
                Text = label,
                Foreground = Brushes.White,
                Background = new SolidColorBrush(Color.FromArgb(128, 0, 0, 0)),
                Padding = new Thickness(2)
            };

            Canvas.SetLeft(textBlock, point.X + 5);
            Canvas.SetTop(textBlock, point.Y - 5);
            _canvas.Children.Add(textBlock);
            _debugElements.Add(textBlock);
        }
    }

    public void ClearDebugElements()
    {
        foreach (var element in _debugElements)
        {
            _canvas.Children.Remove(element);
        }
        _debugElements.Clear();
    }
}