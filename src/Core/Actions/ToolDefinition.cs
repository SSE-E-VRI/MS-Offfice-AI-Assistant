using System;
using System.Collections.Generic;

namespace MSOfficeAIAssistant.Core
{
    /// <summary>
    /// Describes one parameter of an Office tool/action.
    /// </summary>
    public class ToolParameterDefinition
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Description { get; set; }
        public bool IsRequired { get; set; }
        public object DefaultValue { get; set; }

        public ToolParameterDefinition()
        {
            Type = "string";
            IsRequired = false;
        }

        public ToolParameterDefinition(string name, string type, string description, bool isRequired = false, object defaultValue = null)
        {
            Name = name;
            Type = type ?? "string";
            Description = description;
            IsRequired = isRequired;
            DefaultValue = defaultValue;
        }
    }

    /// <summary>
    /// Formal schema definition for an Office host tool/action per SSOT §5.2 and §5.3.
    /// </summary>
    public class ToolDefinition
    {
        public string Name { get; set; }
        public string Host { get; set; }
        public string Description { get; set; }
        public int RiskLevel { get; set; }
        public bool RequiresApproval { get; set; }
        public bool IsUndoable { get; set; }
        public List<ToolParameterDefinition> Parameters { get; set; }
        public List<string> Aliases { get; set; }

        public ToolDefinition()
        {
            Parameters = new List<ToolParameterDefinition>();
            Aliases = new List<string>();
            RiskLevel = 1;
            RequiresApproval = true;
            IsUndoable = true;
        }

        public ToolDefinition(string name, string host, string description, int riskLevel, bool requiresApproval, bool isUndoable)
        {
            Name = name;
            Host = host;
            Description = description;
            RiskLevel = riskLevel;
            RequiresApproval = requiresApproval;
            IsUndoable = isUndoable;
            Parameters = new List<ToolParameterDefinition>();
            Aliases = new List<string>();
        }

        public ToolDefinition WithParameter(string name, string type, string description, bool isRequired = false, object defaultValue = null)
        {
            Parameters.Add(new ToolParameterDefinition(name, type, description, isRequired, defaultValue));
            return this;
        }

        public ToolDefinition WithAlias(string alias)
        {
            if (!string.IsNullOrEmpty(alias) && !Aliases.Contains(alias))
            {
                Aliases.Add(alias);
            }
            return this;
        }

        /// <summary>
        /// Converts this tool definition into standard OpenAI-compatible function schema dictionary.
        /// </summary>
        public Dictionary<string, object> ToOpenAiFunctionSchema()
        {
            var properties = new Dictionary<string, object>();
            var required = new List<string>();

            foreach (var p in Parameters)
            {
                var propDict = new Dictionary<string, object>
                {
                    { "type", p.Type },
                    { "description", p.Description ?? string.Empty }
                };
                if (p.DefaultValue != null)
                {
                    propDict["default"] = p.DefaultValue;
                }
                properties[p.Name] = propDict;

                if (p.IsRequired)
                {
                    required.Add(p.Name);
                }
            }

            var parametersDict = new Dictionary<string, object>
            {
                { "type", "object" },
                { "properties", properties }
            };
            if (required.Count > 0)
            {
                parametersDict["required"] = required;
            }

            var functionDict = new Dictionary<string, object>
            {
                { "name", Name.Replace('.', '_') },
                { "description", Description ?? string.Empty },
                { "parameters", parametersDict }
            };

            return new Dictionary<string, object>
            {
                { "type", "function" },
                { "function", functionDict }
            };
        }
    }
}
