using System;
using System.Collections.Generic;
using System.ComponentModel;
using Newtonsoft.Json;
using MSOfficeAIAssistant.Core.Planning;

namespace MSOfficeAIAssistant.API.Models
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
                    OnPropertyChanged("HasVariants");
                }
            }
        }

        /// <summary>
        /// Marker a "3 Variants" ribbon prompt is instructed to emit between each alternative
        /// rewrite. Kept in one place so the prompt text (RibbonCallback.OnRewriteVariants) and
        /// the parser (RewriteVariantParser) can never drift apart.
        /// </summary>
        public const string VariantDelimiter = "---VARIANT---";

        /// <summary>Computed: true once streaming has produced the variant delimiter, so the
        /// "Compare Variants" button can appear as soon as it's usable rather than only when
        /// streaming finishes.</summary>
        [JsonIgnore]
        public bool HasVariants
        {
            get { return !string.IsNullOrEmpty(_content) && _content.IndexOf(VariantDelimiter, StringComparison.Ordinal) >= 0; }
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

        /// <summary>Full prompt content sent to the API, which may differ from the displayed Content.</summary>
        [JsonProperty("fullContent", NullValueHandling = NullValueHandling.Ignore)]
        public string FullContent { get; set; }

        /// <summary>
        /// Name of the Word bookmark pinning the selection this exchange was generated from
        /// (Selection scope only, Word host only). Lets the Insert button replace exactly the
        /// text that was sent to the model, even if the user has since clicked or scrolled
        /// elsewhere in the document while the response was streaming in. Session-only: a
        /// bookmark does not survive reopening the document, so this is never serialized to the
        /// conversation history.
        /// </summary>
        [JsonIgnore]
        public string SourceSelectionBookmark { get; set; }

        /// <summary>
        /// The un-augmented prompt text and title this response was generated from (before
        /// selection/document context was appended), set only for ribbon-triggered quick prompts
        /// (Rewrite, Tone presets, 3 Variants, Generate, Continue, ...). Lets Regenerate re-run
        /// the same command with freshly read context rather than resending stale prompt text.
        /// Session-only, never serialized to conversation history.
        /// </summary>
        [JsonIgnore]
        public string RegenerateSourcePrompt { get; set; }

        [JsonIgnore]
        public string RegenerateSourceTitle { get; set; }

        /// <summary>Computed: true when this response can be regenerated or discarded.</summary>
        [JsonIgnore]
        public bool CanRegenerate
        {
            get { return !string.IsNullOrEmpty(RegenerateSourcePrompt); }
        }

        private System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.Actions.OfficeAction> _officeActions;
        private System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.SpreadsheetAction> _actions;
        private System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.PowerPointAction> _powerPointActions;
        private bool _hydrated;

        /// <summary>
        /// Single authoritative structured action collection for UI rendering and execution (SSOT §5.3).
        /// </summary>
        [JsonProperty("officeActions", NullValueHandling = NullValueHandling.Ignore)]
        public System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.Actions.OfficeAction> OfficeActions
        {
            get
            {
                EnsureHydrated();
                return _officeActions;
            }
            set
            {
                _officeActions = value;
                OnPropertyChanged("OfficeActions");
                OnPropertyChanged("HasOfficeActions");
            }
        }

        [JsonIgnore]
        public bool HasOfficeActions
        {
            get
            {
                EnsureHydrated();
                return _officeActions != null && _officeActions.Count > 0;
            }
        }

        public void NotifyOfficeActionsChanged()
        {
            OnPropertyChanged("OfficeActions");
            OnPropertyChanged("HasOfficeActions");
        }

        private Plan _plan;

        /// <summary>
        /// Live, in-memory Plan object for Plan-mode message processing.
        /// Not serialized to conversation history (marked JsonIgnore).
        /// </summary>
        [JsonIgnore]
        public Plan Plan
        {
            get { return _plan; }
            set
            {
                if (_plan != value)
                {
                    _plan = value;
                    OnPropertyChanged("Plan");
                    OnPropertyChanged("HasPlan");
                }
            }
        }

        /// <summary>
        /// Computed: true if this message has a Plan.
        /// </summary>
        [JsonIgnore]
        public bool HasPlan
        {
            get { return Plan != null; }
        }

        /// <summary>
        /// Notifies observers that Plan or HasPlan has changed.
        /// Mirrors NotifyOfficeActionsChanged pattern.
        /// </summary>
        public void NotifyPlanChanged()
        {
            OnPropertyChanged("Plan");
            OnPropertyChanged("HasPlan");
        }

        [JsonProperty("actions", NullValueHandling = NullValueHandling.Ignore)]
        public System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.SpreadsheetAction> Actions
        {
            get { return _actions; }
            set
            {
                _actions = value;
                _hydrated = false;
                OnPropertyChanged("Actions");
                OnPropertyChanged("HasActions");
            }
        }

        [JsonIgnore]
        public bool HasActions
        {
            get { return _actions != null && _actions.Count > 0; }
        }

        [JsonProperty("powerPointActions", NullValueHandling = NullValueHandling.Ignore)]
        public System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.PowerPointAction> PowerPointActions
        {
            get { return _powerPointActions; }
            set
            {
                _powerPointActions = value;
                _hydrated = false;
                OnPropertyChanged("PowerPointActions");
                OnPropertyChanged("HasPowerPointActions");
            }
        }

        [JsonIgnore]
        public bool HasPowerPointActions
        {
            get { return _powerPointActions != null && _powerPointActions.Count > 0; }
        }

        /// <summary>
        /// Converts legacy Actions and PowerPointActions from existing conversation stores
        /// into the single OfficeActions collection on deserialization/access.
        /// </summary>
        public void EnsureHydrated()
        {
            if (_hydrated) return;
            _hydrated = true;

            if (_officeActions == null)
            {
                _officeActions = new System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.Actions.OfficeAction>();
            }

            if (_officeActions.Count == 0 && _actions != null && _actions.Count > 0)
            {
                foreach (var act in _actions)
                {
                    var oa = MSOfficeAIAssistant.Core.Actions.OfficeAction.FromSpreadsheetAction(act);
                    if (oa != null)
                    {
                        _officeActions.Add(oa);
                    }
                }
            }

            if (_officeActions.Count == 0 && _powerPointActions != null && _powerPointActions.Count > 0)
            {
                foreach (var act in _powerPointActions)
                {
                    var oa = MSOfficeAIAssistant.Core.Actions.OfficeAction.FromPowerPointAction(act);
                    if (oa != null)
                    {
                        _officeActions.Add(oa);
                    }
                }
            }
        }

        [System.Runtime.Serialization.OnDeserialized]
        internal void OnDeserializedMethod(System.Runtime.Serialization.StreamingContext context)
        {
            EnsureHydrated();
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
            _officeActions = new System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.Actions.OfficeAction>();
            _actions = new System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.SpreadsheetAction>();
            _powerPointActions = new System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.PowerPointAction>();
            _timestamp = DateTime.Now;
        }

        public ChatMessage(string role, string content)
        {
            Role = role;
            Content = content;
            _officeActions = new System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.Actions.OfficeAction>();
            _actions = new System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.SpreadsheetAction>();
            _powerPointActions = new System.Collections.ObjectModel.ObservableCollection<MSOfficeAIAssistant.Core.PowerPointAction>();
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
