using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace Excel2Json.App;

internal sealed class MessageBox : Window
{
    public MessageBox(string title, string message)
    {
        Title = title;
        Width = 420;
        Height = 200;
        Background = Brushes.Transparent;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;

        Content = new Border
        {
            Padding = new Thickness(16),
            Background = new SolidColorBrush(Color.Parse("#16213C")),
            CornerRadius = new CornerRadius(10),
            Child = new StackPanel
            {
                Spacing = 12,
                Children =
                {
                    new TextBlock { Text = message, Foreground = Brushes.White, TextWrapping = TextWrapping.Wrap },
                    new Button
                    {
                        Content = "确定",
                        HorizontalAlignment = HorizontalAlignment.Right,
                        Width = 80
                    }
                }
            }
        };

        if (Content is Border border && border.Child is StackPanel panel && panel.Children.LastOrDefault() is Button ok)
        {
            ok.Click += (_, _) => Close();
        }
    }
}
