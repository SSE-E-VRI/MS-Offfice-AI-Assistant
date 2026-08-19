using System;
using System.Collections.Generic;
using System.ComponentModel;
using Newtonsoft.Json;

namespace MistralOfficeAddin.API.Models
{
    public class ChatMessage : INotifyPropertyChanged
    {
        private string _role;
        private string _content;
        private bool _isStreaming;
        private DateTime _timestamp;

        [JsonProperty("role")]
        public string Role
        {
            get { return _role; }
            set
            {
                if (_role != value)
                {
                    _role = value;
                    OnPropertyChanged("Role");
                    OnPropertyChanged("IsUser");
                    OnPropertyChanged("IsAssistant");
                    OnPropertyChanged("IsSystem");
                }
            }
        }

        [JsonProperty("content")]
        public string Content
        {
            get { return _content; }
            set
            {
                if (_content != value)
                {
                    _content = value;
                    OnPropertyChanged("Content");
                }
            }
        }

        [JsonIgnore]
        public bool IsStreaming
        {
            get { return _isStreaming; }
            set
            {
                if (_isStreaming != value)
                {
                    _isStreaming = value;
                    OnPropertyChanged("IsStreaming");
                }
            }
        }

        [JsonProperty("timestamp", NullValueHandling = NullValueHandling.Ignore)]
        public DateTime Timestamp
        {
            get { return _timestamp; }
            set
            {
                if (_timestamp != value)
                {
                    _timestamp = value;
                    OnPropertyChanged("Timestamp");
                }
            }
        }

        [JsonIgnore]
        public bool IsUser
        {
            get { return string.Equals(Role, "user", StringComparison.OrdinalIgnoreCase); }
        }

        [JsonIgnore]
        public bool IsAssistant
        {
            get { return string.Equals(Role, "assistant", StringComparison.OrdinalIgnoreCase); }
        }

        [JsonIgnore]
        public bool IsSystem
        {
            get { return string.Equals(Role, "system", StringComparison.OrdinalIgnoreCase); }
        }

        private System.Collections.ObjectModel.ObservableCollection<MistralOfficeAddin.Core.SpreadsheetAction> _actions;

        [JsonProperty("actions", NullValueHandling = NullValueHandling.Ignore)]
        public System.Collections.ObjectModel.ObservableCollection<MistralOfficeAddin.Core.SpreadsheetAction> Actions
        {
            get { return _actions; }
            set { _actions = value; OnPropertyChanged("Actions"); OnPropertyChanged("HasActions"); }
        }

        [JsonIgnore]
        public bool HasActions
        {
            get { return _actions != null && _actions.Count > 0; }
        }

        public void NotifyActionsChanged()
        {
            OnPropertyChanged("Actions");
            OnPropertyChanged("HasActions");
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public ChatMessage()
        {
            _actions = new System.Collections.ObjectModel.ObservableCollection<MistralOfficeAddin.Core.SpreadsheetAction>();
            _timestamp = DateTime.Now;
        }

        public ChatMessage(string role, string content)
        {
            Role = role;
            Content = content;
            _actions = new System.Collections.ObjectModel.ObservableCollection<MistralOfficeAddin.Core.SpreadsheetAction>();
            _timestamp = DateTime.Now;
        }
    }

    public class ChatRequest
    {
        [JsonProperty("model")]
        public string Model { get; set; }

        [JsonProperty("messages")]
        public List<ChatMessage> Messages { get; set; }

        [JsonProperty("temperature")]
        public double Temperature { get; set; }

        [JsonProperty("max_tokens")]
        public int MaxTokens { get; set; }

        [JsonProperty("stream")]
        public bool Stream { get; set; }

        [JsonProperty("top_p", NullValueHandling = NullValueHandling.Ignore)]
        public double? TopP { get; set; }

        [JsonProperty("safe_prompt", NullValueHandling = NullValueHandling.Ignore)]
        public bool? SafePrompt { get; set; }

        public ChatRequest()
        {
            Messages = new List<ChatMessage>();
            Temperature = 0.7;
            MaxTokens = 4096;
            Stream = true;
        }
    }
}
