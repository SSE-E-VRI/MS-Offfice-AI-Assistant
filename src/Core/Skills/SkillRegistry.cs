using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Core.Skills
{
    public static class SkillRegistry
    {
        public static List<Skill> LoadPack(string packName)
        {
            if (string.IsNullOrWhiteSpace(packName))
            {
                Logger.Warn("SkillRegistry.LoadPack called with null or empty packName");
                return new List<Skill>();
            }

            try
            {
                string normalizedPack = packName.ToLowerInvariant().Trim();
                string resourceName = string.Format("MSOfficeAIAssistant.Core.Skills.Manifests.{0}.json", normalizedPack);

                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        Logger.Warn(string.Format("SkillRegistry: Could not find embedded resource '{0}'", resourceName));
                        return new List<Skill>();
                    }

                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string json = reader.ReadToEnd();
                        List<Skill> skills = JsonConvert.DeserializeObject<List<Skill>>(json);

                        if (skills == null)
                        {
                            Logger.Warn(string.Format("SkillRegistry: Deserialization of {0} returned null", resourceName));
                            return new List<Skill>();
                        }

                        return skills;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("SkillRegistry.LoadPack failed for pack '{0}': {1}", packName, ex.Message));
                return new List<Skill>();
            }
        }

        public static List<Skill> GetAllPacks()
        {
            var allSkills = new List<Skill>();
            allSkills.AddRange(LoadPack("general"));
            allSkills.AddRange(LoadPack("railway"));
            return allSkills;
        }
    }
}
