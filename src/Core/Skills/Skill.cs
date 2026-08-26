using System.Collections.Generic;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Core.Skills
{
    public class Skill
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("required_context")]
        public List<string> RequiredContext { get; set; }

        [JsonProperty("preferred_host")]
        public string PreferredHost { get; set; }

        [JsonProperty("prompt_template")]
        public string PromptTemplate { get; set; }

        [JsonProperty("output_structure")]
        public string OutputStructure { get; set; }

        [JsonProperty("default_mode")]
        public string DefaultMode { get; set; }

        [JsonProperty("risk_ceiling")]
        public int RiskCeiling { get; set; }

        [JsonProperty("domain_pack")]
        public string DomainPack { get; set; }

        public Skill()
        {
            RequiredContext = new List<string>();
        }
    }
}
