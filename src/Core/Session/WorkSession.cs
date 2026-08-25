using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using MSOfficeAIAssistant.Core.Planning;

namespace MSOfficeAIAssistant.Core.Session
{
    /// <summary>
    /// Represents a persistent work session linking a document, a plan, and execution context.
    /// A WorkSession captures the entire state needed to resume an interrupted multi-step task.
    /// </summary>
    public class WorkSession
    {
        /// <summary>
        /// Unique identifier for this work session (Guid.NewGuid().ToString()).
        /// </summary>
        [JsonProperty("work_session_id")]
        public string WorkSessionId { get; set; }

        /// <summary>
        /// Document key matching ConversationStore's per-document identity convention.
        /// Used to look up the conversation history and context for this session.
        /// </summary>
        [JsonProperty("document_key")]
        public string DocumentKey { get; set; }

        /// <summary>
        /// Human-readable title for this work session, e.g. derived from the original request.
        /// </summary>
        [JsonProperty("title")]
        public string Title { get; set; }

        /// <summary>
        /// The plan associated with this session, or null if no plan has been created yet.
        /// A WorkSession may exist in the early stages before a plan is structured.
        /// </summary>
        [JsonProperty("plan", NullValueHandling = NullValueHandling.Ignore)]
        public Plan Plan { get; set; }

        /// <summary>
        /// Hosts this session has touched, e.g. ["Excel", "Word"].
        /// Used by Phase D3 to detect which host a paused session is waiting to resume on.
        /// </summary>
        [JsonProperty("source_hosts")]
        public List<string> SourceHosts { get; set; }

        /// <summary>
        /// UTC timestamp when this session was created.
        /// </summary>
        [JsonProperty("created_utc")]
        public DateTime CreatedUtc { get; set; }

        /// <summary>
        /// UTC timestamp when this session was last updated.
        /// </summary>
        [JsonProperty("updated_utc")]
        public DateTime UpdatedUtc { get; set; }

        /// <summary>
        /// Free-text status label reflecting the session's current state, e.g. "Draft" | "Awaiting host: Word" | "Completed".
        /// Mirrors Plan.Status when a plan exists. Kept as a string (not an enum) to allow Phase D3 to extend it without modification.
        /// </summary>
        [JsonProperty("status")]
        public string Status { get; set; }

        /// <summary>
        /// Initializes a new WorkSession with a generated WorkSessionId.
        /// </summary>
        public WorkSession()
        {
            WorkSessionId = Guid.NewGuid().ToString();
            SourceHosts = new List<string>();
            CreatedUtc = DateTime.UtcNow;
            UpdatedUtc = DateTime.UtcNow;
            Status = "Draft";
        }
    }
}
