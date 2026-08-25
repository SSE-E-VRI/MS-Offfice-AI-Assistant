using System;
using System.ComponentModel;
using MSOfficeAIAssistant.Core.Actions;

namespace MSOfficeAIAssistant.Core.Planning
{
    public enum PlanStepStatus
    {
        Pending,
        Approved,
        Skipped,
        Applying,
        Applied,
        Failed,
        RolledBack
    }

    /// <summary>
    /// A single step in an executable Plan - either an OfficeAction mutation or a reasoning-only narrative step.
    /// </summary>
    public class PlanStep : INotifyPropertyChanged
    {
        private int _order;
        private string _description;
        private PlanStepStatus _status;
        private string _resultText;
        private string _errorMessage;

        /// <summary>
        /// Unique identifier for this step.
        /// </summary>
        public string StepId { get; set; }

        /// <summary>
        /// 1-based position in the plan, mutable (user can reorder before execution).
        /// </summary>
        public int Order
        {
            get { return _order; }
            set
            {
                if (_order != value)
                {
                    _order = value;
                    OnPropertyChanged("Order");
                }
            }
        }

        /// <summary>
        /// Human-readable summary of this step.
        /// </summary>
        public string Description
        {
            get { return _description; }
            set
            {
                if (_description != value)
                {
                    _description = value;
                    OnPropertyChanged("Description");
                }
            }
        }

        /// <summary>
        /// The Office action for this step, or null if this is a reasoning-only step.
        /// </summary>
        public OfficeAction Action { get; set; }

        /// <summary>
        /// Computed: true when Action == null (reasoning/narrative step with no Office mutation).
        /// </summary>
        public bool IsReasoningOnly
        {
            get { return Action == null; }
        }

        /// <summary>
        /// Target host: "Word" | "Excel" | "PowerPoint".
        /// For action steps, mirrors Action.Host; for reasoning steps, specifies the context.
        /// </summary>
        public string TargetHost { get; set; }

        /// <summary>
        /// Current status of the step.
        /// </summary>
        public PlanStepStatus Status
        {
            get { return _status; }
            set
            {
                if (_status != value)
                {
                    _status = value;
                    OnPropertyChanged("Status");
                    OnPropertyChanged("StatusDisplay");
                }
            }
        }

        /// <summary>
        /// Human-readable status display string, mirroring the pattern in SpreadsheetAction.cs.
        /// </summary>
        public string StatusDisplay
        {
            get
            {
                switch (_status)
                {
                    case PlanStepStatus.Applied:
                        return string.IsNullOrEmpty(_resultText) ? "✓ Applied" : string.Format("✓ Applied ({0})", _resultText);
                    case PlanStepStatus.Failed:
                        return string.IsNullOrEmpty(_errorMessage) ? "⚠ Failed" : string.Format("⚠ Failed: {0}", _errorMessage);
                    case PlanStepStatus.RolledBack:
                        return "⟲ Rolled Back";
                    case PlanStepStatus.Applying:
                        return "⟳ Applying...";
                    case PlanStepStatus.Approved:
                        return "✓ Approved";
                    case PlanStepStatus.Skipped:
                        return "⊘ Skipped";
                    default:
                        return "○ Pending";
                }
            }
        }

        /// <summary>
        /// Result text from execution, if any.
        /// </summary>
        public string ResultText
        {
            get { return _resultText; }
            set
            {
                if (_resultText != value)
                {
                    _resultText = value;
                    OnPropertyChanged("ResultText");
                    OnPropertyChanged("StatusDisplay");
                }
            }
        }

        /// <summary>
        /// Error message if the step failed.
        /// </summary>
        public string ErrorMessage
        {
            get { return _errorMessage; }
            set
            {
                if (_errorMessage != value)
                {
                    _errorMessage = value;
                    OnPropertyChanged("ErrorMessage");
                    OnPropertyChanged("StatusDisplay");
                }
            }
        }

        /// <summary>
        /// Initializes a new PlanStep with a generated StepId.
        /// </summary>
        public PlanStep()
        {
            StepId = Guid.NewGuid().ToString();
            Status = PlanStepStatus.Pending;
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
    }
}
