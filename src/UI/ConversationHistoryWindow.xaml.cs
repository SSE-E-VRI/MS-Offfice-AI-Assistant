using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.UI
{
    public partial class ConversationHistoryWindow : Window
    {
        private List<ConversationSessionSummary> _sessions;
        public string SelectedDocumentKey { get; private set; }

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

        public ConversationHistoryWindow(List<ConversationSessionSummary> sessions = null)
        {
            EnsureWpfApplication();
            InitializeComponent();

            _sessions = sessions ?? ConversationStore.Instance.ListSessions();
            RefreshSessionsList();
        }

        private void RefreshSessionsList()
        {
            if (_sessions == null || _sessions.Count == 0)
            {
                SessionsListBox.Visibility = Visibility.Collapsed;
                EmptyStateText.Visibility = Visibility.Visible;
            }
            else
            {
                SessionsListBox.ItemsSource = _sessions;
                SessionsListBox.Visibility = Visibility.Visible;
                EmptyStateText.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnLoad_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;

            var item = btn.DataContext as ConversationSessionSummary;
            if (item == null) return;

            SelectedDocumentKey = item.DocumentKey;
            this.DialogResult = true;
            this.Close();
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;

            var item = btn.DataContext as ConversationSessionSummary;
            if (item == null) return;

            string message = string.Format("Delete conversation '{0}'? This cannot be undone.", item.Title);
            if (MessageBox.Show(message, "Confirm Deletion", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }

            try
            {
                ConversationStore.Instance.ClearHistory(item.DocumentKey);
                _sessions.Remove(item);
                RefreshSessionsList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("Could not delete conversation: {0}", ex.Message),
                    "Deletion Failed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
