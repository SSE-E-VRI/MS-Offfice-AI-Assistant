using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using MSOfficeAIAssistant.API.Models;

namespace MSOfficeAIAssistant.API
{
    public static class TokenCounter
    {
        private static readonly Regex WordRegex = new Regex(@"\S+", RegexOptions.Compiled);

        public static int CountTokens(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;

            // Character heuristic: ~4 characters per token on average for English/code
            int charBased = (int)Math.Ceiling(text.Length / 3.8);

            // Word heuristic: ~1.3 tokens per word
            int wordCount = WordRegex.Matches(text).Count;
            int wordBased = (int)Math.Ceiling(wordCount * 1.33);

            // Blended estimation
            return Math.Max(charBased, wordBased);
        }

        public static List<ChatMessage> TruncateToFit(List<ChatMessage> messages, int maxTokensLimit, string systemPrompt = null)
        {
            if (messages == null || messages.Count == 0)
            {
                var empty = new List<ChatMessage>();
                if (!string.IsNullOrEmpty(systemPrompt))
                {
                    empty.Add(new ChatMessage("system", systemPrompt));
                }
                return empty;
            }

            var result = new List<ChatMessage>();
            int reservedTokens = !string.IsNullOrEmpty(systemPrompt) ? CountTokens(systemPrompt) + 10 : 0;
            int availableTokens = Math.Max(500, maxTokensLimit - reservedTokens);

            // Walk backwards from latest messages
            var reversed = new List<ChatMessage>();
            int currentTokens = 0;

            for (int i = messages.Count - 1; i >= 0; i--)
            {
                var msg = messages[i];
                if (msg.IsSystem) continue; // handle system separately

                int msgTokens = CountTokens(msg.Content) + 4;
                if (currentTokens + msgTokens > availableTokens)
                {
                    // If even the latest user message exceeds available tokens, truncate its content
                    if (reversed.Count == 0 && msg.IsUser)
                    {
                        int allowedChars = Math.Max(100, availableTokens * 3);
                        string truncatedContent = msg.Content.Length > allowedChars
                            ? msg.Content.Substring(0, allowedChars) + "\n...[truncated for token limit]"
                            : msg.Content;
                        reversed.Add(new ChatMessage(msg.Role, truncatedContent));
                    }
                    break;
                }

                reversed.Add(msg);
                currentTokens += msgTokens;
            }

            reversed.Reverse();

            if (!string.IsNullOrEmpty(systemPrompt))
            {
                result.Add(new ChatMessage("system", systemPrompt));
            }

            result.AddRange(reversed);
            return result;
        }
    }
}
