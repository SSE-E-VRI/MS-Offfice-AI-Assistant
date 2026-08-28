using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.Integration;
using System.Windows.Media;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.UI.Helpers;

namespace MSOfficeAIAssistant.UI
{
    /// <summary>
    /// Carousel for a "3 Variants" rewrite response (RibbonCallback.OnRewriteVariants,
    /// RewriteVariantParser.Split). Lets the user page through the alternatives and either
    /// insert the one showing, ask for a fresh set (Regenerate), or walk away (Discard) without
    /// touching the document. Modeled on ResponsePreviewWindow's code-only WPF Window pattern,
    /// but this one returns a result to the caller via callbacks instead of being fire-and-forget.
    /// </summary>
    public class RewriteVariantsWindow : Window
    {
        private static RewriteVariantsWindow _openWindow;

        private readonly List<string> _variants;
        private readonly Action<string> _onInsert;
        private readonly Action _onRegenerate;
        private readonly Action _onDiscard;
        private int _index;

        private readonly TextBlock _formattedView;
        private TextBlock _counterBlock;
        private Button _prevButton;
        private Button _nextButton;

        private static readonly SolidColorBrush InkStrong = Frozen(Color.FromRgb(0x0F, 0x17, 0x2A));
        private static readonly SolidColorBrush Ink = Frozen(Color.FromRgb(0x1E, 0x29, 0x3B));
        private static readonly SolidColorBrush InkMuted = Frozen(Color.FromRgb(0x47, 0x55, 0x69));
        private static readonly SolidColorBrush Border0 = Frozen(Color.FromRgb(0xE2, 0xE8, 0xF0));
        private static readonly SolidColorBrush SurfaceAlt = Frozen(Color.FromRgb(0xF8, 0xFA, 0xFC));
        private static readonly SolidColorBrush PrimaryStrong = Frozen(Color.FromRgb(0x1D, 0x4E, 0xD8));

        private static SolidColorBrush Frozen(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }

        /// <summary>
        /// Opens the carousel for the given variants, reusing the existing window if one is
        /// already up. No-ops if fewer than 1 variant was parsed.
        /// </summary>
        public static void ShowFor(List<string> variants, Action<string> onInsert, Action onRegenerate, Action onDiscard)
        {
            if (variants == null || variants.Count == 0) return;

            try
            {
                if (_openWindow != null)
                {
                    _openWindow.Close();
                    _openWindow = null;
                }

                var window = new RewriteVariantsWindow(variants, onInsert, onRegenerate, onDiscard);
                _openWindow = window;
                window.Closed += delegate { if (ReferenceEquals(_openWindow, window)) _openWindow = null; };
                window.Show();
                window.Activate();
            }
            catch (Exception ex)
            {
                Logger.Error("RewriteVariantsWindow could not be opened", ex);
            }
        }

        private RewriteVariantsWindow(List<string> variants, Action<string> onInsert, Action onRegenerate, Action onDiscard)
        {
            _variants = variants;
            _onInsert = onInsert;
            _onRegenerate = onRegenerate;
            _onDiscard = onDiscard;

            this.Title = "Rewrite Variants";
            this.Width = 780;
            this.Height = 640;
            this.MinWidth = 420;
            this.MinHeight = 320;
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.Background = Brushes.White;
            this.FontFamily = new FontFamily("Segoe UI");
            this.FontSize = 12;

            try
            {
                ElementHost.EnableModelessKeyboardInterop(this);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("RewriteVariantsWindow keyboard interop setup failed: {0}", ex.Message));
            }

            _formattedView = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                Foreground = Ink,
                FontSize = 12.5,
                LineHeight = 18,
                LineStackingStrategy = LineStackingStrategy.BlockLineHeight
            };

            var scroller = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(16, 12, 16, 12),
                Content = _formattedView
            };

            var body = new Border
            {
                BorderBrush = Border0,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Background = Brushes.White,
                Margin = new Thickness(12, 0, 12, 12),
                Child = scroller
            };

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var header = BuildHeader();
            Grid.SetRow(header, 0);

            var nav = BuildNav();
            Grid.SetRow(nav, 1);

            Grid.SetRow(body, 2);

            var footer = BuildFooter();
            Grid.SetRow(footer, 3);

            grid.Children.Add(header);
            grid.Children.Add(nav);
            grid.Children.Add(body);
            grid.Children.Add(footer);
            this.Content = grid;

            RenderCurrent();
        }

        private UIElement BuildHeader()
        {
            var stack = new StackPanel { Margin = new Thickness(12, 12, 12, 4) };
            stack.Children.Add(new TextBlock
            {
                Text = "Rewrite Variants",
                FontSize = 15,
                FontWeight = FontWeights.Bold,
                Foreground = InkStrong
            });
            stack.Children.Add(new TextBlock
            {
                Foreground = InkMuted,
                FontSize = 11,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 2, 0, 0),
                Text = "Page through the alternatives below, then Insert the one you want."
            });
            return stack;
        }

        private UIElement BuildNav()
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(12, 4, 12, 8)
            };

            _prevButton = new Button
            {
                Content = "< Prev",
                Width = 70,
                Padding = new Thickness(6, 3, 6, 3),
                Background = SurfaceAlt,
                Foreground = Ink,
                BorderBrush = Border0,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            _prevButton.Click += delegate { MoveTo(_index - 1); };

            var counter = new TextBlock
            {
                Foreground = InkMuted,
                FontSize = 12,
                Margin = new Thickness(12, 0, 12, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Width = 70,
                TextAlignment = TextAlignment.Center
            };
            _counterBlock = counter;

            _nextButton = new Button
            {
                Content = "Next >",
                Width = 70,
                Padding = new Thickness(6, 3, 6, 3),
                Background = SurfaceAlt,
                Foreground = Ink,
                BorderBrush = Border0,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            _nextButton.Click += delegate { MoveTo(_index + 1); };

            panel.Children.Add(_prevButton);
            panel.Children.Add(counter);
            panel.Children.Add(_nextButton);
            return panel;
        }

        private void MoveTo(int index)
        {
            if (index < 0 || index >= _variants.Count) return;
            _index = index;
            RenderCurrent();
        }

        private void RenderCurrent()
        {
            MarkdownHelper.SetMarkdownText(_formattedView, _variants[_index]);
            if (_counterBlock != null)
                _counterBlock.Text = string.Format("{0} of {1}", _index + 1, _variants.Count);
            if (_prevButton != null) _prevButton.IsEnabled = _index > 0;
            if (_nextButton != null) _nextButton.IsEnabled = _index < _variants.Count - 1;
        }

        private UIElement BuildFooter()
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(12, 0, 12, 12)
            };

            var discard = new Button
            {
                Content = "Discard",
                Width = 85,
                Padding = new Thickness(8, 4, 8, 4),
                Margin = new Thickness(0, 0, 8, 0),
                Background = SurfaceAlt,
                Foreground = Ink,
                BorderBrush = Border0,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            discard.Click += delegate
            {
                this.Close();
                if (_onDiscard != null) _onDiscard();
            };

            var regenerate = new Button
            {
                Content = "Regenerate",
                Width = 95,
                Padding = new Thickness(8, 4, 8, 4),
                Margin = new Thickness(0, 0, 8, 0),
                Background = SurfaceAlt,
                Foreground = Ink,
                BorderBrush = Border0,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            regenerate.Click += delegate
            {
                this.Close();
                if (_onRegenerate != null) _onRegenerate();
            };

            var insert = new Button
            {
                Content = "Insert",
                Width = 85,
                Padding = new Thickness(8, 4, 8, 4),
                Background = PrimaryStrong,
                Foreground = Brushes.White,
                BorderBrush = Border0,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            insert.Click += delegate
            {
                string chosen = _variants[_index];
                this.Close();
                if (_onInsert != null) _onInsert(chosen);
            };

            panel.Children.Add(discard);
            panel.Children.Add(regenerate);
            panel.Children.Add(insert);
            return panel;
        }
    }
}
