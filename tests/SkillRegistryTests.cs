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
            TestAllSkillsHaveRequiredFields();
        }

        private static void TestLoadGeneralPack()
        {
            var skills = SkillRegistry.LoadPack("general");
            Assert(skills != null, "LoadPack should never return null");
            Assert(skills.Count == 9, "general pack should have exactly 9 skills");

            bool foundOfficialLetter = false;
            foreach (var skill in skills)
            {
                Assert(skill.DomainPack == "general", "All general pack skills should have DomainPack 'general'");
                if (skill.Id == "official_letter")
                    foundOfficialLetter = true;
            }
            Assert(foundOfficialLetter, "general pack should contain 'official_letter' skill");
        }

        private static void TestLoadRailwayPack()
        {
            var skills = SkillRegistry.LoadPack("railway");
            Assert(skills != null, "LoadPack should never return null");
            Assert(skills.Count == 13, "railway pack should have exactly 13 skills");

            bool foundOfficialLetter = false;
            bool foundDrmBriefing = false;
            foreach (var skill in skills)
            {
                Assert(skill.DomainPack == "railway", "All railway pack skills should have DomainPack 'railway'");
                if (skill.Id == "official_letter")
                    foundOfficialLetter = true;
                if (skill.Id == "drm_briefing")
                    foundDrmBriefing = true;
            }
            Assert(foundOfficialLetter, "railway pack should contain 'official_letter' skill (general duplicated in railway)");
            Assert(foundDrmBriefing, "railway pack should contain 'drm_briefing' railway-specific skill");
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
            Assert(allSkills.Count == 22, "GetAllPacks should return 22 skills total (9 general + 13 railway)");

            int generalCount = 0;
            int railwayCount = 0;
            foreach (var skill in allSkills)
            {
                if (skill.DomainPack == "general") generalCount++;
                if (skill.DomainPack == "railway") railwayCount++;
            }
            Assert(generalCount == 9, "GetAllPacks should include exactly 9 general skills");
            Assert(railwayCount == 13, "GetAllPacks should include exactly 13 railway skills");
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

        private static void TestAllSkillsHaveRequiredFields()
        {
            var allSkills = SkillRegistry.GetAllPacks();
            Assert(allSkills.Count > 0, "Should have skills loaded");

            foreach (var skill in allSkills)
            {
                Assert(!string.IsNullOrWhiteSpace(skill.Id), "Skill must have non-empty Id");
                Assert(!string.IsNullOrWhiteSpace(skill.Name), "Skill must have non-empty Name");
                Assert(!string.IsNullOrWhiteSpace(skill.Description), "Skill must have non-empty Description");
                Assert(!string.IsNullOrWhiteSpace(skill.PromptTemplate), "Skill must have non-empty PromptTemplate");
            }
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
