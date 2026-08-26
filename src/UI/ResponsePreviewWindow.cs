using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.Integration;
using System.Windows.Media;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.UI.Helpers;

namespace MSOfficeAIAssistant.UI
{
    /// <summary>
    /// Preview of an AI response as it will land in the document.
    ///
    /// This replaces a raw MessageBox.Show(msg.Content), which showed the response as unformatted
    /// text: literal "**" markers, and the model's trailing commentary ("Key Features", "Structure")
    /// that the Insert button no longer applies. The window renders the cleaned content with the
    /// same Markdown renderer the chat pane uses, so what is previewed matches what Insert writes,
    /// and keeps a "Show raw response" toggle for the unedited model output.
    /// </summary>
    public class ResponsePreviewWindow : Window
    {
        private static ResponsePreviewWindow _openWindow;

        private readonly string _rawResponse;
        private readonly string _cleanedResponse;
        private readonly TextBlock _formattedView;
        private readonly TextBox _rawView;
        private readonly ScrollViewer _scroller;
        // Not readonly: unlike the other fields, this is built inside BuildHeader(), a
        // helper called FROM the constructor rather than assigned directly in its body --
        // C# only allows a readonly field to be set at that outer level (CS0191).
        private CheckBox _showRaw;

        private static readonly SolidColorBrush InkStrong = Frozen(Color.FromRgb(0x0F, 0x17, 0x2A));
        private static readonly SolidColorBrush Ink = Frozen(Color.FromRgb(0x1E, 0x29, 0x3B));
        private static readonly SolidColorBrush InkMuted = Frozen(Color.FromRgb(0x47, 0x55, 0x69));
        private static readonly SolidColorBrush Border0 = Frozen(Color.FromRgb(0xE2, 0xE8, 0xF0));
        private static readonly SolidColorBrush SurfaceAlt = Frozen(Color.FromRgb(0xF8, 0xFA, 0xFC));
        private static readonly SolidColorBrush Warning = Frozen(Color.FromRgb(0xB4, 0x53, 0x09));

        private static SolidColorBrush Frozen(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }

        /// <summary>Opens the preview, reusing the existing window if one is already up.</summary>
        public static void ShowFor(string response)
        {
            if (string.IsNullOrWhiteSpace(response)) return;

            try
            {
                if (_openWindow != null)
                {
                    _openWindow.Close();
                    _openWindow = null;
                }

                var window = new ResponsePreviewWindow(response);
                _openWindow = window;
                window.Closed += delegate { if (ReferenceEquals(_openWindow, window)) _openWindow = null; };
                window.Show();
                window.Activate();
            }
            catch (Exception ex)
            {
                Logger.Error("ResponsePreviewWindow could not be opened", ex);
                // Losing the preview should never lose the response.
                MessageBox.Show(response, "AI Response Preview", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        /// <summary>The unedited model output, shown by the "Show raw response" toggle.</summary>
        public string RawText { get { return _rawResponse; } }

        /// <summary>The content actually rendered, and the content Insert will apply.</summary>
        public string PreviewText { get { return _cleanedResponse; } }

        /// <summary>Public for tests; use ShowFor to open the preview normally.</summary>
        public ResponsePreviewWindow(string response)
        {
            _rawResponse = response ?? string.Empty;
            _cleanedResponse = CleanForPreview(_rawResponse);

            this.Title = "AI Response Preview";
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
                // Same interop the Help and chat windows use: without it a modeless WPF window
                // gets no keyboard input while Office owns the message loop.
                ElementHost.EnableModelessKeyboardInterop(this);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ResponsePreviewWindow keyboard interop setup failed: {0}", ex.Message));
            }

            _formattedView = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                Foreground = Ink,
                FontSize = 12.5,
                LineHeight = 18,
                LineStackingStrategy = LineStackingStrategy.BlockLineHeight
            };
            MarkdownHelper.SetMarkdownText(_formattedView, _cleanedResponse);

            _rawView = new TextBox
            {
                Text = _rawResponse,
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                AcceptsReturn = true,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Foreground = Ink,
                FontFamily = new FontFamily("Consolas"),
                FontSize = 12,
                Visibility = Visibility.Collapsed
            };

            var content = new Grid();
            content.Children.Add(_formattedView);
            content.Children.Add(_rawView);

            _scroller = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(16, 12, 16, 12),
                Content = content
            };

            var body = new Border
            {
                BorderBrush = Border0,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Background = Brushes.White,
                Margin = new Thickness(12, 0, 12, 12),
                Child = _scroller
            };

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var header = BuildHeader();
            Grid.SetRow(header, 0);
            Grid.SetRow(body, 1);

            var footer = BuildFooter();
            Grid.SetRow(footer, 2);

            grid.Children.Add(header);
            grid.Children.Add(body);
            grid.Children.Add(footer);
            this.Content = grid;
        }

        private UIElement BuildHeader()
        {
            var stack = new StackPanel { Margin = new Thickness(12, 12, 12, 8) };

            stack.Children.Add(new TextBlock
            {
                Text = "AI Response Preview",
                FontSize = 15,
                FontWeight = FontWeights.Bold,
                Foreground = InkStrong
            });

            _showRaw = new CheckBox
            {
                Content = "Show raw response",
                Foreground = InkMuted,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 6, 0, 0)
            };
            _showRaw.Checked += delegate { SetRawVisible(true); };
            _showRaw.Unchecked += delegate { SetRawVisible(false); };

            var caption = new TextBlock
            {
                Foreground = InkMuted,
                FontSize = 11,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 2, 0, 0),
                Text = WasCleaned()
                    ? "This is what Insert will put in the document. The model's lead-in and closing notes are not included."
                    : "This is what Insert will put in the document."
            };
            stack.Children.Add(caption);

            // A response can read as an analysis OF the text -- a comparison table, a "Key
            // Improvements" rundown, a second "recommended" draft -- rather than a single
            // replacement. Insert applies the whole thing verbatim, table included, so this is
            // called out before that happens rather than after.
            if (ResponseContentCleaner.LooksLikeEditAnalysisReport(_cleanedResponse))
            {
                stack.Children.Add(new TextBlock
                {
                    Foreground = Warning,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 6, 0, 0),
                    Text = "\u26a0 This reads like an analysis of the text (a comparison table and/or a list of " +
                           "changes and rationale) rather than a single replacement. Insert will apply all of it, " +
                           "table included, in place of your selection."
                });
            }

            stack.Children.Add(_showRaw);
            return stack;
        }

        private UIElement BuildFooter()
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(12, 0, 12, 12)
            };

            var copy = new Button
            {
                Content = "Copy",
                Width = 85,
                Padding = new Thickness(8, 4, 8, 4),
                Margin = new Thickness(0, 0, 8, 0),
                Background = SurfaceAlt,
                Foreground = Ink,
                BorderBrush = Border0,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            copy.Click += delegate { MarkdownClipboard.SetResponse(_rawResponse); };

            var close = new Button
            {
                Content = "Close",
                Width = 85,
                Padding = new Thickness(8, 4, 8, 4),
                Background = SurfaceAlt,
                Foreground = Ink,
                BorderBrush = Border0,
                IsCancel = true,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            close.Click += delegate { this.Close(); };

            panel.Children.Add(copy);
            panel.Children.Add(close);
            return panel;
        }

        private void SetRawVisible(bool raw)
        {
            _rawView.Visibility = raw ? Visibility.Visible : Visibility.Collapsed;
            _formattedView.Visibility = raw ? Visibility.Collapsed : Visibility.Visible;
            try { _scroller.ScrollToTop(); }
            catch { }
        }

        private bool WasCleaned()
        {
            return !string.Equals(_rawResponse.Trim(), _cleanedResponse.Trim(), StringComparison.Ordinal);
        }

        private static string CleanForPreview(string response)
        {
            try
            {
                bool strip = ConfigManager.Instance == null || ConfigManager.Instance.StripConversationalWrapper;
                return strip ? ResponseContentCleaner.ExtractInsertableContent(response) : response;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ResponsePreviewWindow: cleaner skipped ({0})", ex.Message));
                return response;
            }
        }
    }
}
