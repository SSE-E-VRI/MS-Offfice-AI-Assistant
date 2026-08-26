using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.Core.Skills;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Tests
{
    public static class SkillRegistryTests
    {
        public static void RunAll()
        {
            TestLoadGeneralPack();
            TestLoadRailwayPack();
            TestCaseInsensitivePackName();
            TestNonexistentPackReturnsEmpty();
            TestGetAllPacksCombinesBoth();
            TestSkillRoundTrip();
        }

        private static void TestLoadGeneralPack()
        {
            var skills = SkillRegistry.LoadPack("general");
            Assert(skills != null, "LoadPack should never return null");
            Assert(skills.Count == 1, "general pack should have exactly 1 placeholder skill");
            Assert(skills[0].Id == "placeholder_general_summary", "First skill ID should be placeholder_general_summary");
            Assert(skills[0].DomainPack == "general", "DomainPack should be 'general'");
        }

        private static void TestLoadRailwayPack()
        {
            var skills = SkillRegistry.LoadPack("railway");
            Assert(skills != null, "LoadPack should never return null");
            Assert(skills.Count == 1, "railway pack should have exactly 1 placeholder skill");
            Assert(skills[0].Id == "placeholder_railway_timetable", "First skill ID should be placeholder_railway_timetable");
            Assert(skills[0].DomainPack == "railway", "DomainPack should be 'railway'");
        }

        private static void TestCaseInsensitivePackName()
        {
            var skillsLower = SkillRegistry.LoadPack("general");
            var skillsUpper = SkillRegistry.LoadPack("GENERAL");
            var skillsMixed = SkillRegistry.LoadPack("General");

            Assert(skillsLower.Count == skillsUpper.Count, "Case should not affect pack loading");
            Assert(skillsLower.Count == skillsMixed.Count, "Case should not affect pack loading");
            Assert(skillsLower[0].Id == skillsUpper[0].Id, "Case-insensitive load should return same skill ID");
        }

        private static void TestNonexistentPackReturnsEmpty()
        {
            var skills = SkillRegistry.LoadPack("nonexistent_pack");
            Assert(skills != null, "LoadPack should return empty list, not null");
            Assert(skills.Count == 0, "Nonexistent pack should return empty list");
        }

        private static void TestGetAllPacksCombinesBoth()
        {
            var allSkills = SkillRegistry.GetAllPacks();
            Assert(allSkills != null, "GetAllPacks should never return null");
            Assert(allSkills.Count == 2, "GetAllPacks should return 2 skills total (1 general + 1 railway)");

            bool hasGeneral = false;
            bool hasRailway = false;
            foreach (var skill in allSkills)
            {
                if (skill.DomainPack == "general") hasGeneral = true;
                if (skill.DomainPack == "railway") hasRailway = true;
            }
            Assert(hasGeneral, "GetAllPacks should include a general skill");
            Assert(hasRailway, "GetAllPacks should include a railway skill");
        }

        private static void TestSkillRoundTrip()
        {
            var skill = new Skill
            {
                Id = "test_skill",
                Name = "Test Skill",
                Description = "A test skill",
                RequiredContext = new List<string> { "Selection", "CurrentFile" },
                PreferredHost = "Word",
                PromptTemplate = "Do something with {content}",
                OutputStructure = "JSON object",
                DefaultMode = "Edit",
                RiskCeiling = 2,
                DomainPack = "general"
            };

            string json = JsonConvert.SerializeObject(skill);
            Skill deserialized = JsonConvert.DeserializeObject<Skill>(json);

            Assert(deserialized != null, "Deserialization should not return null");
            Assert(deserialized.Id == "test_skill", "Id should round-trip correctly");
            Assert(deserialized.Name == "Test Skill", "Name should round-trip correctly");
            Assert(deserialized.Description == "A test skill", "Description should round-trip correctly");
            Assert(deserialized.RequiredContext.Count == 2, "RequiredContext should round-trip correctly");
            Assert(deserialized.RequiredContext[0] == "Selection", "RequiredContext[0] should be 'Selection'");
            Assert(deserialized.RequiredContext[1] == "CurrentFile", "RequiredContext[1] should be 'CurrentFile'");
            Assert(deserialized.PreferredHost == "Word", "PreferredHost should round-trip correctly");
            Assert(deserialized.PromptTemplate == "Do something with {content}", "PromptTemplate should round-trip correctly");
            Assert(deserialized.OutputStructure == "JSON object", "OutputStructure should round-trip correctly");
            Assert(deserialized.DefaultMode == "Edit", "DefaultMode should round-trip correctly");
            Assert(deserialized.RiskCeiling == 2, "RiskCeiling should round-trip correctly");
            Assert(deserialized.DomainPack == "general", "DomainPack should round-trip correctly");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion failed: " + message);
            }
        }
    }
}
