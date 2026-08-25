using System;
using System.Collections.Generic;
using System.IO;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.Core.Planning;
using MSOfficeAIAssistant.Core.Session;
using MSOfficeAIAssistant.Core.Actions;

namespace MSOfficeAIAssistant.Tests
{
    public static class WorkSessionStoreTests
    {
        public static void RunAll()
        {
            TestRoundTripWithPlan();
            TestRoundTripPreservesPlanStepAndActionStatus();
            TestRoundTripWithNullPlan();
            TestListByDocumentKeyOrdering();
            TestDeleteSession();
            TestCorruptedFileHandling();
        }

        private static void TestRoundTripPreservesPlanStepAndActionStatus()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "test-worksession-status-" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(tempDir);
                var store = new WorkSessionStore(tempDir);

                var plan = new Plan();
                plan.Status = PlanStatus.Executing;

                var appliedAction = new OfficeAction();
                appliedAction.ActionId = Guid.NewGuid().ToString();
                appliedAction.Host = "Excel";
                appliedAction.Operation = "excel.write_formula";
                appliedAction.Target = new ActionTarget { Range = "B2" };
                appliedAction.Status = OfficeActionStatus.Applied;
                appliedAction.Rollback = new RollbackInfo("mock_success");
                appliedAction.Rollback.IsRollbackPossible = true;

                var appliedStep = new PlanStep();
                appliedStep.Order = 1;
                appliedStep.Description = "Applied mutation";
                appliedStep.TargetHost = "Excel";
                appliedStep.Status = PlanStepStatus.Applied;
                appliedStep.Action = appliedAction;
                plan.Steps.Add(appliedStep);

                var failedStep = new PlanStep();
                failedStep.Order = 2;
                failedStep.Description = "Failed step";
                failedStep.TargetHost = "Word";
                failedStep.Status = PlanStepStatus.Failed;
                failedStep.ErrorMessage = "boom";
                plan.Steps.Add(failedStep);

                var session = new WorkSession();
                session.DocumentKey = "StatusRoundTripDoc";
                session.Title = "Status round-trip";
                session.Plan = plan;
                session.Status = "Failed";
                session.SourceHosts = new List<string> { "Excel", "Word" };

                store.Save(session);
                var loaded = store.Load(session.WorkSessionId);

                Assert(loaded != null && loaded.Plan != null, "Session and Plan loaded");
                Assert(loaded.Plan.Status == PlanStatus.Executing, "PlanStatus enum round-trips");
                Assert(loaded.Plan.Steps[0].Status == PlanStepStatus.Applied, "PlanStep Applied round-trips");
                Assert(loaded.Plan.Steps[1].Status == PlanStepStatus.Failed, "PlanStep Failed round-trips");
                Assert(loaded.Plan.Steps[1].ErrorMessage == "boom", "PlanStep ErrorMessage round-trips");
                Assert(loaded.Plan.Steps[0].Action != null, "Action present after load");
                Assert(loaded.Plan.Steps[0].Action.Status == OfficeActionStatus.Applied,
                    "OfficeAction.Status must round-trip (required for RollbackAll after resume)");
                Assert(loaded.Plan.Steps[0].Action.Rollback != null
                    && loaded.Plan.Steps[0].Action.Rollback.IsRollbackPossible,
                    "RollbackInfo survives round-trip");
            }
            finally
            {
                try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
            }
        }

        private static void TestRoundTripWithPlan()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "test-worksession-" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(tempDir);
                var store = new WorkSessionStore(tempDir);

                // Create a plan with 2 PlanSteps
                var plan = new Plan();
                plan.Title = "Test Plan";
                plan.SourceRequest = "Test request";
                plan.Status = PlanStatus.Draft;

                // Step 1: Reasoning-only step (no Action)
                var reasoningStep = new PlanStep();
                reasoningStep.Order = 1;
                reasoningStep.Description = "Analyze the requirements";
                reasoningStep.TargetHost = "Excel";
                // Action is null for reasoning-only step
                plan.Steps.Add(reasoningStep);

                // Step 2: Action step
                var actionStep = new PlanStep();
                actionStep.Order = 2;
                actionStep.Description = "Insert formulas";
                actionStep.TargetHost = "Excel";
                actionStep.Action = new OfficeAction();
                actionStep.Action.ActionId = Guid.NewGuid().ToString();
                actionStep.Action.Host = "Excel";
                actionStep.Action.Operation = "SetFormula";
                actionStep.Action.Target = new ActionTarget { Range = "B2:B10" };
                actionStep.Action.SourceReason = "Insert SUM formula";
                actionStep.Action.Status = OfficeActionStatus.Pending;
                plan.Steps.Add(actionStep);

                // Create a WorkSession with the plan
                var session = new WorkSession();
                session.WorkSessionId = Guid.NewGuid().ToString();
                session.DocumentKey = "TestDocument123";
                session.Title = "Test Work Session";
                session.Plan = plan;
                session.SourceHosts = new List<string> { "Excel" };
                session.Status = "Draft";

                // Save
                store.Save(session);

                // Load
                var loaded = store.Load(session.WorkSessionId);

                // Verify all fields match
                Assert(loaded != null, "Session loaded successfully");
                Assert(loaded.WorkSessionId == session.WorkSessionId, "WorkSessionId matches");
                Assert(loaded.DocumentKey == session.DocumentKey, "DocumentKey matches");
                Assert(loaded.Title == session.Title, "Title matches");
                Assert(loaded.Status == session.Status, "Status matches");
                Assert(loaded.SourceHosts != null && loaded.SourceHosts.Count == 1, "SourceHosts count matches");
                Assert(loaded.SourceHosts[0] == "Excel", "SourceHosts[0] is Excel");

                // Verify Plan details
                Assert(loaded.Plan != null, "Plan is not null");
                Assert(loaded.Plan.Title == plan.Title, "Plan.Title matches");
                Assert(loaded.Plan.SourceRequest == plan.SourceRequest, "Plan.SourceRequest matches");
                Assert(loaded.Plan.Steps.Count == 2, "Plan has 2 steps");

                // Verify reasoning step
                Assert(loaded.Plan.Steps[0].Order == 1, "First step Order is 1");
                Assert(loaded.Plan.Steps[0].Description == "Analyze the requirements", "First step Description matches");
                Assert(loaded.Plan.Steps[0].Action == null, "First step Action is null");
                Assert(loaded.Plan.Steps[0].IsReasoningOnly, "First step IsReasoningOnly is true");

                // Verify action step
                Assert(loaded.Plan.Steps[1].Order == 2, "Second step Order is 2");
                Assert(loaded.Plan.Steps[1].Description == "Insert formulas", "Second step Description matches");
                Assert(loaded.Plan.Steps[1].Action != null, "Second step Action is not null");
                Assert(loaded.Plan.Steps[1].Action.Operation == "SetFormula", "Action.Operation matches");
                Assert(loaded.Plan.Steps[1].Action.TargetDisplay == "B2:B10", "Action.TargetDisplay matches");
                Assert(!loaded.Plan.Steps[1].IsReasoningOnly, "Second step IsReasoningOnly is false");
            }
            finally
            {
                try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
            }
        }

        private static void TestRoundTripWithNullPlan()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "test-worksession-" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(tempDir);
                var store = new WorkSessionStore(tempDir);

                var session = new WorkSession();
                session.WorkSessionId = Guid.NewGuid().ToString();
                session.DocumentKey = "DocumentNoPlan";
                session.Title = "Session Without Plan";
                session.Plan = null;  // No plan yet
                session.SourceHosts = new List<string>();
                session.Status = "Initial";

                // Save and load
                store.Save(session);
                var loaded = store.Load(session.WorkSessionId);

                // Verify
                Assert(loaded != null, "Session loaded successfully");
                Assert(loaded.Plan == null, "Plan is null as expected");
                Assert(loaded.Title == session.Title, "Title matches");
                Assert(loaded.DocumentKey == session.DocumentKey, "DocumentKey matches");
            }
            finally
            {
                try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
            }
        }

        private static void TestListByDocumentKeyOrdering()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "test-worksession-" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(tempDir);
                var store = new WorkSessionStore(tempDir);

                // Create sessions for different documents
                var session1 = new WorkSession();
                session1.DocumentKey = "Doc1";
                session1.Title = "Doc1 Session 1";
                store.Save(session1);
                System.Threading.Thread.Sleep(10);  // Small delay to ensure different timestamps

                var session2 = new WorkSession();
                session2.DocumentKey = "Doc1";
                session2.Title = "Doc1 Session 2";
                store.Save(session2);
                System.Threading.Thread.Sleep(10);

                var session3 = new WorkSession();
                session3.DocumentKey = "Doc2";
                session3.Title = "Doc2 Session 1";
                store.Save(session3);

                // List sessions for Doc1
                var doc1Sessions = store.ListByDocumentKey("Doc1");
                Assert(doc1Sessions.Count == 2, "Doc1 has 2 sessions");
                Assert(doc1Sessions[0].WorkSessionId == session2.WorkSessionId, "Most recent Doc1 session is first (session2)");
                Assert(doc1Sessions[1].WorkSessionId == session1.WorkSessionId, "Older Doc1 session is second (session1)");

                // List sessions for Doc2
                var doc2Sessions = store.ListByDocumentKey("Doc2");
                Assert(doc2Sessions.Count == 1, "Doc2 has 1 session");
                Assert(doc2Sessions[0].WorkSessionId == session3.WorkSessionId, "Doc2 session matches");
            }
            finally
            {
                try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
            }
        }

        private static void TestDeleteSession()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "test-worksession-" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(tempDir);
                var store = new WorkSessionStore(tempDir);

                var session = new WorkSession();
                session.DocumentKey = "DocDelete";
                store.Save(session);

                var loaded1 = store.Load(session.WorkSessionId);
                Assert(loaded1 != null, "Session loaded before delete");

                // Delete
                store.Delete(session.WorkSessionId);

                // Try to load again
                var loaded2 = store.Load(session.WorkSessionId);
                Assert(loaded2 == null, "Session is null after delete");

                // Verify ListByDocumentKey no longer returns this session
                var sessions = store.ListByDocumentKey("DocDelete");
                Assert(sessions.Count == 0, "Session no longer appears in ListByDocumentKey");
            }
            finally
            {
                try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
            }
        }

        private static void TestCorruptedFileHandling()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "test-worksession-" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(tempDir);
                var store = new WorkSessionStore(tempDir);

                // Create and save a valid session
                var session = new WorkSession();
                session.DocumentKey = "DocCorrupt";
                session.Title = "Will be corrupted";
                store.Save(session);

                string sessionFile = Path.Combine(tempDir, session.WorkSessionId + ".dat");
                Assert(File.Exists(sessionFile), "Session file exists");

                // Corrupt the file by overwriting with garbage
                File.WriteAllBytes(sessionFile, new byte[] { 0xFF, 0xFE, 0xFD, 0x00, 0x00 });

                // Try to load - should return null without throwing
                WorkSession loaded = null;
                Exception loadException = null;
                try
                {
                    loaded = store.Load(session.WorkSessionId);
                }
                catch (Exception ex)
                {
                    loadException = ex;
                }

                // Verify behavior
                Assert(loadException == null, "Load did not throw an exception");
                Assert(loaded == null, "Corrupted session returns null");

                // Verify the corrupt file was quarantined to .bak
                string bakFile = sessionFile + ".bak";
                Assert(File.Exists(bakFile), "Corrupted file was quarantined to .bak");
                Assert(!File.Exists(sessionFile), "Original corrupt file was moved (not present)");

                // Verify the session no longer appears in ListByDocumentKey
                var sessions = store.ListByDocumentKey("DocCorrupt");
                Assert(sessions.Count == 0, "Corrupted session does not appear in list");
            }
            finally
            {
                try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion Failed: " + message);
            }
        }
    }
}
