using System;
using System.Collections.Generic;
using System.Windows;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.UI
{
    public partial class ActionHistoryWindow : Window
    {
        private static void EnsureWpfApplication()
        {
            if (System.Windows.Application.Current != null) return;
            try
            {
                var app = new System.Windows.Application();
                app.ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown;
            }
            catch { }
        }

        public ActionHistoryWindow(List<ActionAuditEntry> entries = null)
        {
            EnsureWpfApplication();
            InitializeComponent();

            if (entries == null || entries.Count == 0)
            {
                ActionsListBox.Visibility = Visibility.Collapsed;
                EmptyStateText.Visibility = Visibility.Visible;
            }
            else
            {
                ActionsListBox.ItemsSource = entries;
                ActionsListBox.Visibility = Visibility.Visible;
                EmptyStateText.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
