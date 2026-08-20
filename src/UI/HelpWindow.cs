using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.Integration;
using System.Windows.Interop;
using MistralOfficeAddin.Core;

namespace MistralOfficeAddin.UI
{
    /// <summary>
    /// In-add-in window that shows the offline User Manual (UserManual.html) embedded as a
    /// manifest resource. Keeps the manual in front of the Office document window and works
    /// without any API key. The HTML is written to a local cache file so in-page TOC anchors
    /// navigate reliably in the IE-based WebBrowser engine.
    /// </summary>
    public class HelpWindow : Window
    {
        private const string ResourceName = "MSOfficeAIAssistant.Help.UserManual.html";
        private static HelpWindow _openWindow;

        private HelpWindow()
        {
            this.Title = "AI Assistant — User Manual";
            this.Width = 920;
            this.Height = 720;
            this.MinWidth = 620;
            this.MinHeight = 460;
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.Background = System.Windows.Media.Brushes.White;
            this.ShowActivated = true;

            try
            {
                // Excel 2010 owns the native message loop; this bridge lets a modeless WPF
                // window receive keyboard input while Excel remains active (same as chat).
                ElementHost.EnableModelessKeyboardInterop(this);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("HelpWindow keyboard interop setup failed: {0}", ex.Message));
            }

            var topBar = new DockPanel { LastChildFill = true, Margin = new Thickness(8, 6, 8, 6) };
            var openInBrowser = new Button
            {
                Content = "Open in browser (print)",
                Width = 150,
                Padding = new Thickness(8, 3, 8, 3),
                Margin = new Thickness(0, 0, 4, 0)
            };
            openInBrowser.Click += delegate { OpenInBrowser(); };
            DockPanel.SetDock(openInBrowser, Dock.Right);

            var hint = new TextBlock
            {
                Text = "AI Assistant — User Manual (v0.4.0)",
                VerticalAlignment = VerticalAlignment.Center,
                FontWeight = FontWeights.SemiBold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(15, 23, 42))
            };
            DockPanel.SetDock(hint, Dock.Left);

            topBar.Children.Add(openInBrowser);
            topBar.Children.Add(hint);

            var browser = new WebBrowser { VerticalAlignment = VerticalAlignment.Stretch };

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            Grid.SetRow(topBar, 0);
            Grid.SetRow(browser, 1);
            grid.Children.Add(topBar);
            grid.Children.Add(browser);

            this.Content = grid;

            this.Closed += delegate { _openWindow = null; };
            LoadUserManual(browser);
        }

        private static void EnsureWpfApplication()
        {
            if (Application.Current != null) return;
            try
            {
                var app = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
            }
            catch { }
        }

        /// <summary>
        /// Opens the manual. If a manual window is already open it is brought to the front
        /// instead of stacking a second copy.
        /// </summary>
        public static void ShowHelp(IntPtr ownerHwnd)
        {
            try
            {
                if (_openWindow != null)
                {
                    try
                    {
                        if (_openWindow.WindowState == WindowState.Minimized)
                            _openWindow.WindowState = WindowState.Normal;
                        _openWindow.Activate();
                        return;
                    }
                    catch { }
                }

                EnsureWpfApplication();

                var win = new HelpWindow();
                if (ownerHwnd != IntPtr.Zero)
                {
                    try { new WindowInteropHelper(win).Owner = ownerHwnd; }
                    catch { }
                }
                _openWindow = win;
                win.Show();
                win.Activate();
            }
            catch (Exception ex)
            {
                Logger.Error("Could not open the User Manual window", ex);
                MessageBox.Show(
                    string.Format("Could not open the User Manual: {0}", ex.Message),
                    "AI Assistant",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void OpenInBrowser()
        {
            try
            {
                string cachePath = WriteHtmlToCacheFile();
                if (!string.IsNullOrWhiteSpace(cachePath) && File.Exists(cachePath))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = cachePath,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("OpenInBrowser failed: {0}", ex.Message));
            }
        }

        private void LoadUserManual(WebBrowser browser)
        {
            string html = null;
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream(ResourceName))
                {
                    if (stream != null)
                    {
                        using (var reader = new StreamReader(stream))
                            html = reader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Could not read embedded User Manual resource", ex);
            }

            if (string.IsNullOrWhiteSpace(html))
            {
                MessageBox.Show(
                    "The User Manual resource is missing from the add-in.\n\n" +
                    "Reinstall the add-in or run install.cmd again, then retry.",
                    "AI Assistant — User Manual",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                this.Close();
                return;
            }

            try
            {
                string cachePath = WriteHtmlToCacheFile(html);
                if (!string.IsNullOrWhiteSpace(cachePath) && File.Exists(cachePath))
                {
                    browser.Navigate(new Uri(cachePath));
                    return;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("HelpWindow cache write failed, falling back to in-memory: {0}", ex.Message));
            }

            // Fallback: render directly from the embedded string. In-page anchor jumps are
            // less reliable here, but the content is fully available.
            try { browser.NavigateToString(html); }
            catch (Exception ex) { Logger.Error("HelpWindow NavigateToString failed", ex); }
        }

        private static string WriteHtmlToCacheFile()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream(ResourceName))
                {
                    if (stream == null) return null;
                    using (var reader = new StreamReader(stream))
                        return WriteHtmlToCacheFile(reader.ReadToEnd());
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("HelpWindow cache write failed: {0}", ex.Message));
                return null;
            }
        }

        private static string WriteHtmlToCacheFile(string html)
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MistralOfficeAddin");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                string path = Path.Combine(dir, "UserManual.html");
                File.WriteAllText(path, html);
                return path;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("HelpWindow cache write failed: {0}", ex.Message));
                return null;
            }
        }
    }
}