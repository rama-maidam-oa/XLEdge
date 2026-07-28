using MahApps.Metro.IconPacks;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using XLEdge.Helpers;
using XLEdge.Utilities;

namespace XLEdge.Views
{
    /// <summary>
    /// Interaction logic for XLEdgeMessageWindow.xaml
    /// </summary>
    public partial class XLEdgeMessageWindow : DpiAwareWindow
    {
        public MessageBoxResult Result { get; private set; } = MessageBoxResult.None;
        public XLEdgeMessageWindow(string message,
                               MessageBoxIcon icon,
                               MessageBoxButtons buttons = MessageBoxButtons.OK)
        {
            InitializeComponent();

            EnhancedDragDropHelper.EnableWindowDrag(this);


            MsgText.Text = message;

            SetMessageIcon(icon);
            SetupButtons(buttons);

            // One-time sizing pass: measures the window's natural size for this message, then
            // locks SizeToContent to Manual so the user can freely resize the window afterward.
            this.Loaded += XLEdgeMessageWindow_AutoFitOnce;
        }

        private void XLEdgeMessageWindow_AutoFitOnce(object sender, RoutedEventArgs e)
        {
            this.Loaded -= XLEdgeMessageWindow_AutoFitOnce;

            try
            {
                this.SizeToContent = SizeToContent.WidthAndHeight;
                this.UpdateLayout();

                double fitWidth = Math.Max(this.MinWidth, Math.Min(this.MaxWidth, this.ActualWidth));
                double fitHeight = Math.Max(this.MinHeight, Math.Min(this.MaxHeight, this.ActualHeight));

                this.SizeToContent = SizeToContent.Manual;
                this.Width = fitWidth;
                this.Height = fitHeight;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "XLEdgeMessageWindow_AutoFitOnce");
                this.SizeToContent = SizeToContent.Manual;
            }
        }
        private void HeaderPanel_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
                DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Result = MessageBoxResult.Cancel;
            DialogResult = false;
            Close();
        }

        private void SetMessageIcon(MessageBoxIcon icon)
        {
            switch (icon)
            {
                case MessageBoxIcon.Error:
                    MsgIcon.Kind = PackIconFontAwesomeKind.CircleXmarkSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(231, 76, 60));
                    break;

                case MessageBoxIcon.Warning:
                    MsgIcon.Kind = PackIconFontAwesomeKind.TriangleExclamationSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(241, 196, 15));
                    break;

                case MessageBoxIcon.Information:
                    MsgIcon.Kind = PackIconFontAwesomeKind.CircleInfoSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(46, 134, 171));
                    break;

                case MessageBoxIcon.Question:
                    MsgIcon.Kind = PackIconFontAwesomeKind.CircleQuestionSolid;
                    MsgIcon.Foreground = new SolidColorBrush(Color.FromRgb(41, 128, 185));
                    break;

                default:
                    MsgIcon.Kind = PackIconFontAwesomeKind.MessageSolid;
                    MsgIcon.Foreground = Brushes.Gray;
                    break;
            }
        }

        // Semantic icon colors: green = proceed (OK/Yes), red = negative/destructive (No), amber = neutral abort (Cancel).
        private static readonly Color GreenIconColor = Color.FromRgb(0x27, 0xAE, 0x60);
        private static readonly Color RedIconColor = Color.FromRgb(0xE7, 0x4C, 0x3C);
        private static readonly Color AmberIconColor = Color.FromRgb(0xF3, 0x9C, 0x12);

        private void SetupButtons(MessageBoxButtons buttons)
        {
            ButtonPanel.Children.Clear();

            switch (buttons)
            {
                case MessageBoxButtons.OK:
                    AddDialogButton("OK", MessageBoxResult.OK, PackIconFontAwesomeKind.CircleCheckSolid, GreenIconColor);
                    break;

                case MessageBoxButtons.OKCancel:
                    AddDialogButton("OK", MessageBoxResult.OK, PackIconFontAwesomeKind.CircleCheckSolid, GreenIconColor);
                    AddDialogButton("Cancel", MessageBoxResult.Cancel, PackIconFontAwesomeKind.CircleXmarkSolid, AmberIconColor);
                    break;

                case MessageBoxButtons.YesNo:
                    AddDialogButton("Yes", MessageBoxResult.Yes, PackIconFontAwesomeKind.CircleCheckSolid, GreenIconColor);
                    AddDialogButton("No", MessageBoxResult.No, PackIconFontAwesomeKind.CircleXmarkSolid, RedIconColor);
                    break;

                case MessageBoxButtons.YesNoCancel:
                    // Each button gets a distinct icon and color so Yes/No/Cancel are visually distinguishable.
                    AddDialogButton("Yes", MessageBoxResult.Yes, PackIconFontAwesomeKind.CircleCheckSolid, GreenIconColor);
                    AddDialogButton("No", MessageBoxResult.No, PackIconFontAwesomeKind.CircleXmarkSolid, RedIconColor);
                    AddDialogButton("Cancel", MessageBoxResult.Cancel, PackIconFontAwesomeKind.CircleMinusSolid, AmberIconColor);
                    break;
            }
        }

        // Builds one dialog button (icon + text) using the shared DynamicContentButton style, sized
        // to fit its content, and wires its Click handler to close the window with the given result.
        private void AddDialogButton(string text, MessageBoxResult resultValue, PackIconFontAwesomeKind iconKind, Color iconColor)
        {
            var btn = new System.Windows.Controls.Button
            {
                Style = (Style)this.FindResource("DynamicContentButton"),
                MinWidth = 90,
                Height = 36,
                Padding = new Thickness(14, 0, 14, 0),
                Margin = new Thickness(ButtonPanel.Children.Count == 0 ? 0 : 8, 0, 0, 0)
            };

            // Build a StackPanel for icon + text
            var panel = new StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center
            };

            // Create the icon - explicit semantic color (green/red/amber) so Yes/No/Cancel are
            // visually distinguishable at a glance, not just by their (also-differentiated) icon shape.
            var faIcon = new PackIconFontAwesome
            {
                Kind = iconKind,
                Width = 16,
                Height = 16,
                Margin = new Thickness(0, 0, 5, 0),
                Foreground = new SolidColorBrush(iconColor)
            };

            // Create the text - no local Foreground, inherits from the Button so it tracks the shared
            // style's rest/hover Foreground correctly.
            var txt = new TextBlock
            {
                Text = text,
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center
            };

            // Add icon + text into panel
            panel.Children.Add(faIcon);
            panel.Children.Add(txt);

            // Set as button content
            btn.Content = panel;

            btn.Click += (sender, e) =>
            {
                Result = resultValue;
                this.DialogResult = true;
                this.Close();
            };

            ButtonPanel.Children.Add(btn);
        }
        private static Color LightenColor(Color color, double amount)
        {
            // Convert to HSL
            RgbToHsl(color, out double h, out double s, out double l);

            // Increase lightness
            l = Math.Max(0.0, Math.Min(1.0, l + amount));

            // Convert back to RGB
            return HslToRgb(h, s, l);
        }
        private static void RgbToHsl(Color color, out double h, out double s, out double l)
        {
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;

            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));

            l = (max + min) / 2.0;

            const double epsilon = 1e-8;
            if (Math.Abs(max - min) < epsilon)
            {
                h = 0;
                s = 0;
            }
            else
            {
                double d = max - min;
                s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);

                if (Math.Abs(max - r) < epsilon)
                    h = (g - b) / d + (g < b ? 6 : 0);
                else if (Math.Abs(max - g) < epsilon)
                    h = (b - r) / d + 2;
                else
                    h = (r - g) / d + 4;

                h /= 6;
            }
        }

        private static Color HslToRgb(double h, double s, double l)
        {
            double r, g, b;
            const double epsilon = 1e-8;

            if (Math.Abs(s) < epsilon)
            {
                r = g = b = l;
            }
            else
            {
                double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
                double p = 2 * l - q;

                r = HueToRgb(p, q, h + 1.0 / 3.0);
                g = HueToRgb(p, q, h);
                b = HueToRgb(p, q, h - 1.0 / 3.0);
            }

            return Color.FromRgb(
                (byte)(r * 255),
                (byte)(g * 255),
                (byte)(b * 255));
        }

        private static double HueToRgb(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1.0 / 6.0) return p + (q - p) * 6 * t;
            if (t < 1.0 / 2.0) return q;
            if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6;
            return p;
        }
    }
}
