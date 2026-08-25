using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace MSOfficeAIAssistant.Core.Planning
{
    public enum PlanStatus
    {
        Draft,
        Approved,
        Executing,
        Paused,
        Completed,
        Failed,
        Cancelled
    }

    /// <summary>
    /// A structured plan containing an ordered sequence of PlanSteps that can be edited before execution.
    /// </summary>
    public class Plan : INotifyPropertyChanged
    {
        private string _planId;
        private string _title;
        private string _sourceRequest;
        private ObservableCollection<PlanStep> _steps;
        private PlanStatus _status;
        private DateTime _createdUtc;
        private DateTime _updatedUtc;

        /// <summary>
        /// Unique identifier for this plan.
        /// </summary>
        public string PlanId
        {
            get { return _planId; }
            set
            {
                if (_planId != value)
                {
                    _planId = value;
                    OnPropertyChanged("PlanId");
                }
            }
        }

        /// <summary>
        /// Short human-readable title, e.g. derived from the user's original request.
        /// </summary>
        public string Title
        {
            get { return _title; }
            set
            {
                if (_title != value)
                {
                    _title = value;
                    OnPropertyChanged("Title");
                }
            }
        }

        /// <summary>
        /// The original user prompt that produced this plan.
        /// </summary>
        public string SourceRequest
        {
            get { return _sourceRequest; }
            set
            {
                if (_sourceRequest != value)
                {
                    _sourceRequest = value;
                    OnPropertyChanged("SourceRequest");
                }
            }
        }

        /// <summary>
        /// Ordered collection of steps in this plan.
        /// </summary>
        public ObservableCollection<PlanStep> Steps
        {
            get { return _steps; }
            set
            {
                if (_steps != value)
                {
                    _steps = value;
                    OnPropertyChanged("Steps");
                }
            }
        }

        /// <summary>
        /// Current status of the plan.
        /// </summary>
        public PlanStatus Status
        {
            get { return _status; }
            set
            {
                if (_status != value)
                {
                    _status = value;
                    OnPropertyChanged("Status");
                }
            }
        }

        /// <summary>
        /// UTC timestamp when this plan was created.
        /// </summary>
        public DateTime CreatedUtc
        {
            get { return _createdUtc; }
            set
            {
                if (_createdUtc != value)
                {
                    _createdUtc = value;
                    OnPropertyChanged("CreatedUtc");
                }
            }
        }

        /// <summary>
        /// UTC timestamp when this plan was last updated.
        /// </summary>
        public DateTime UpdatedUtc
        {
            get { return _updatedUtc; }
            set
            {
                if (_updatedUtc != value)
                {
                    _updatedUtc = value;
                    OnPropertyChanged("UpdatedUtc");
                }
            }
        }

        /// <summary>
        /// Computed: true if the plan has at least one step.
        /// </summary>
        public bool HasSteps
        {
            get { return Steps != null && Steps.Count > 0; }
        }

        /// <summary>
        /// Computed: total number of steps in the plan.
        /// </summary>
        public int TotalStepCount
        {
            get { return Steps != null ? Steps.Count : 0; }
        }

        /// <summary>
        /// Computed: number of steps with Status == Applied.
        /// </summary>
        public int CompletedStepCount
        {
            get
            {
                if (Steps == null) return 0;
                return Steps.Count(s => s.Status == PlanStepStatus.Applied);
            }
        }

        /// <summary>
        /// Computed: ordered list of distinct TargetHost values across all steps.
        /// Used by Phase D3 to detect multi-host plans.
        /// </summary>
        public List<string> DistinctHosts
        {
            get
            {
                if (Steps == null || Steps.Count == 0) return new List<string>();
                var hosts = new List<string>();
                foreach (var step in Steps)
                {
                    if (!string.IsNullOrEmpty(step.TargetHost) && !hosts.Contains(step.TargetHost))
                    {
                        hosts.Add(step.TargetHost);
                    }
                }
                return hosts;
            }
        }

        /// <summary>
        /// Initializes a new Plan with generated PlanId.
        /// </summary>
        public Plan()
        {
            PlanId = Guid.NewGuid().ToString();
            Status = PlanStatus.Draft;
            CreatedUtc = DateTime.UtcNow;
            UpdatedUtc = DateTime.UtcNow;
            Steps = new ObservableCollection<PlanStep>();
        }

        /// <summary>
        /// Moves a step up by one position, swapping Order with the previous step.
        /// </summary>
        public void MoveStepUp(string stepId)
        {
            if (Steps == null || Steps.Count < 2) return;

            int index = -1;
            for (int i = 0; i < Steps.Count; i++)
            {
                if (Steps[i].StepId == stepId)
                {
                    index = i;
                    break;
                }
            }

            if (index <= 0) return; // Already at top or not found

            // Swap Order values
            int tempOrder = Steps[index - 1].Order;
            Steps[index - 1].Order = Steps[index].Order;
            Steps[index].Order = tempOrder;

            // Move in collection
            Steps.Move(index, index - 1);
            UpdatedUtc = DateTime.UtcNow;
        }

        /// <summary>
        /// Moves a step down by one position, swapping Order with the next step.
        /// </summary>
        public void MoveStepDown(string stepId)
        {
            if (Steps == null || Steps.Count < 2) return;

            int index = -1;
            for (int i = 0; i < Steps.Count; i++)
            {
                if (Steps[i].StepId == stepId)
                {
                    index = i;
                    break;
                }
            }

            if (index < 0 || index >= Steps.Count - 1) return; // At bottom or not found

            // Swap Order values
            int tempOrder = Steps[index + 1].Order;
            Steps[index + 1].Order = Steps[index].Order;
            Steps[index].Order = tempOrder;

            // Move in collection
            Steps.Move(index, index + 1);
            UpdatedUtc = DateTime.UtcNow;
        }

        /// <summary>
        /// Removes a step from the plan by StepId.
        /// </summary>
        public void RemoveStep(string stepId)
        {
            if (Steps == null) return;

            int index = -1;
            for (int i = 0; i < Steps.Count; i++)
            {
                if (Steps[i].StepId == stepId)
                {
                    index = i;
                    break;
                }
            }

            if (index >= 0)
            {
                Steps.RemoveAt(index);
                RenumberSteps();
                UpdatedUtc = DateTime.UtcNow;
            }
        }

        /// <summary>
        /// Reassigns Order values to 1..N after any edit, keeping Steps sorted by Order.
        /// </summary>
        public void RenumberSteps()
        {
            if (Steps == null) return;

            // Sort by current Order
            var sorted = Steps.OrderBy(s => s.Order).ToList();

            // Reassign Order 1..N
            for (int i = 0; i < sorted.Count; i++)
            {
                sorted[i].Order = i + 1;
            }

            // Rebuild Steps collection in sorted order
            Steps.Clear();
            foreach (var step in sorted)
            {
                Steps.Add(step);
            }

            UpdatedUtc = DateTime.UtcNow;
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
