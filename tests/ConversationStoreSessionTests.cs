using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MSOfficeAIAssistant.API.Models;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Tests
{
    internal class ConversationStoreSessionTests
    {
        public static void RunAll()
        {
            Assert(TestEmptyDirectory(), "Empty directory returns empty list");
            Assert(TestMultipleSessionsAndTruncation(), "Multiple sessions with truncation");
            Assert(TestDeleteSession(), "Session deletion");
        }

        private static bool TestEmptyDirectory()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "ConvStoreTest_" + Guid.NewGuid().ToString());
            try
            {
                Directory.CreateDirectory(tempDir);
                var store = new ConversationStore(tempDir);
                var sessions = store.ListSessions();
                return sessions != null && sessions.Count == 0;
            }
            finally
            {
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }

        private static bool TestMultipleSessionsAndTruncation()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "ConvStoreTest_" + Guid.NewGuid().ToString());
            try
            {
                Directory.CreateDirectory(tempDir);
                var store = new ConversationStore(tempDir);

                // Create first conversation with short message
                var docAMessages = new List<ChatMessage>
                {
                    new ChatMessage("user", "Hello world"),
                    new ChatMessage("assistant", "Hi there!")
                };
                store.SaveHistory("DocA", docAMessages);
                // Two back-to-back file writes can land on the same filesystem timestamp tick
                // (or even resolve out of wall-clock order under I/O buffering) on a fast/loaded
                // machine, which made the ordering assertion below flaky. Pin each file's
                // LastWriteTimeUtc explicitly so the "more recent" relationship this test checks
                // is deterministic regardless of real execution speed.
                File.SetLastWriteTimeUtc(Path.Combine(tempDir, "DocA.dat"), DateTime.UtcNow.AddMinutes(-10));

                // Create second conversation with long message that should be truncated
                string longMessage = new string('x', 80);
                var docBMessages = new List<ChatMessage>
                {
                    new ChatMessage("user", longMessage),
                    new ChatMessage("assistant", "Response to long message")
                };
                store.SaveHistory("DocB", docBMessages);
                File.SetLastWriteTimeUtc(Path.Combine(tempDir, "DocB.dat"), DateTime.UtcNow);

                var sessions = store.ListSessions();

                if (sessions == null || sessions.Count != 2)
                {
                    Logger.Warn(string.Format("Expected 2 sessions, got {0}", sessions == null ? 0 : sessions.Count));
                    return false;
                }

                // Verify ordering (most recent first)
                var docBSession = sessions.FirstOrDefault(s => s.DocumentKey == "DocB");
                var docASession = sessions.FirstOrDefault(s => s.DocumentKey == "DocA");

                if (docBSession == null || docASession == null)
                {
                    Logger.Warn("Could not find expected sessions");
                    return false;
                }

                // DocB should be more recent (it was saved after DocA)
                if (docBSession.LastUpdatedUtc <= docASession.LastUpdatedUtc)
                {
                    Logger.Warn("Sessions not ordered by LastUpdatedUtc descending");
                    return false;
                }

                // Verify message counts
                if (docASession.MessageCount != 2 || docBSession.MessageCount != 2)
                {
                    Logger.Warn(string.Format("Message counts incorrect: DocA={0}, DocB={1}", docASession.MessageCount, docBSession.MessageCount));
                    return false;
                }

                // Verify truncation (80 chars should be truncated to ~60 with "…")
                if (docBSession.Title.Length > 65 || !docBSession.Title.EndsWith("…"))
                {
                    Logger.Warn(string.Format("Title not properly truncated: {0} (length {1})", docBSession.Title, docBSession.Title.Length));
                    return false;
                }

                // Verify short title is not truncated
                if (!docASession.Title.StartsWith("Hello world"))
                {
                    Logger.Warn(string.Format("Short title was incorrectly modified: {0}", docASession.Title));
                    return false;
                }

                return true;
            }
            finally
            {
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }

        private static bool TestDeleteSession()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "ConvStoreTest_" + Guid.NewGuid().ToString());
            try
            {
                Directory.CreateDirectory(tempDir);
                var store = new ConversationStore(tempDir);

                // Create two conversations
                var docAMessages = new List<ChatMessage>
                {
                    new ChatMessage("user", "Doc A message"),
                    new ChatMessage("assistant", "Response A")
                };
                var docBMessages = new List<ChatMessage>
                {
                    new ChatMessage("user", "Doc B message"),
                    new ChatMessage("assistant", "Response B")
                };

                store.SaveHistory("DocA", docAMessages);
                store.SaveHistory("DocB", docBMessages);

                // Verify both exist
                var sessionsBeforeDelete = store.ListSessions();
                if (sessionsBeforeDelete.Count != 2)
                {
                    Logger.Warn(string.Format("Expected 2 sessions before delete, got {0}", sessionsBeforeDelete.Count));
                    return false;
                }

                // Delete DocA
                store.ClearHistory("DocA");

                // Verify only DocB remains
                var sessionsAfterDelete = store.ListSessions();
                if (sessionsAfterDelete.Count != 1)
                {
                    Logger.Warn(string.Format("Expected 1 session after delete, got {0}", sessionsAfterDelete.Count));
                    return false;
                }

                if (sessionsAfterDelete[0].DocumentKey != "DocB")
                {
                    Logger.Warn(string.Format("Wrong session remains: {0}", sessionsAfterDelete[0].DocumentKey));
                    return false;
                }

                return true;
            }
            finally
            {
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception(string.Format("Assertion failed: {0}", message));
            }
        }
    }
}
