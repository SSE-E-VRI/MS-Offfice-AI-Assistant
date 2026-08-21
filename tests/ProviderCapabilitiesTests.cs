using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.API.Models;
using MSOfficeAIAssistant.Providers;
using Newtonsoft.Json.Linq;

namespace MSOfficeAIAssistant.Tests
{
    public static class ProviderCapabilitiesTests
    {
        public static void RunAll()
        {
            TestCapabilityFlagValues();
            TestProviderDeclaredCapabilities();
            TestOpenAICompatibleBuildPayloadStandard();
            TestOpenAICompatibleBuildPayloadJsonObject();
            TestOpenAICompatibleBuildPayloadToolsAndChoice();
            TestOpenAICompatibleBuildPayloadExtraParameters();
        }

        private static void TestCapabilityFlagValues()
        {
            Assert((int)AICapabilities.StructuredOutput == 32, "StructuredOutput must equal 32");
            Assert((int)AICapabilities.ToolCalling == 64, "ToolCalling must equal 64");
            Assert((int)AICapabilities.JsonMode == 128, "JsonMode must equal 128");

            var combined = AICapabilities.Chat | AICapabilities.StructuredOutput | AICapabilities.ToolCalling | AICapabilities.JsonMode;
            Assert((combined & AICapabilities.StructuredOutput) != 0, "Flag check failed for StructuredOutput");
            Assert((combined & AICapabilities.ToolCalling) != 0, "Flag check failed for ToolCalling");
            Assert((combined & AICapabilities.JsonMode) != 0, "Flag check failed for JsonMode");
        }

        private static void TestProviderDeclaredCapabilities()
        {
            using (var mistral = new MistralProvider("https://api.mistral.ai/v1", "dummy_key"))
            {
                Assert((mistral.Capabilities & AICapabilities.StructuredOutput) != 0, "Mistral must declare StructuredOutput");
                Assert((mistral.Capabilities & AICapabilities.ToolCalling) != 0, "Mistral must declare ToolCalling");
                Assert((mistral.Capabilities & AICapabilities.JsonMode) != 0, "Mistral must declare JsonMode");
            }

            using (var gemini = new GeminiProvider("dummy_key"))
            {
                Assert((gemini.Capabilities & AICapabilities.StructuredOutput) != 0, "Gemini must declare StructuredOutput");
                Assert((gemini.Capabilities & AICapabilities.ToolCalling) != 0, "Gemini must declare ToolCalling");
                Assert((gemini.Capabilities & AICapabilities.JsonMode) != 0, "Gemini must declare JsonMode");
            }

            using (var groq = new GroqProvider("dummy_key"))
            {
                Assert((groq.Capabilities & AICapabilities.ToolCalling) != 0, "Groq must declare ToolCalling");
                Assert((groq.Capabilities & AICapabilities.JsonMode) != 0, "Groq must declare JsonMode");
            }

            using (var custom = new CustomApiProvider("http://localhost:11434/v1", "dummy_key"))
            {
                Assert((custom.Capabilities & AICapabilities.JsonMode) != 0, "Custom provider must declare JsonMode");
            }
        }

        private static void TestOpenAICompatibleBuildPayloadStandard()
        {
            var req = new AIRequest
            {
                Model = "test-model",
                Messages = new List<ChatMessage> { new ChatMessage("user", "Hello world") },
                Temperature = 0.5,
                MaxTokens = 1000
            };

            var payload = OpenAICompatibleClient.BuildPayload(req, stream: false);
            var json = JObject.FromObject(payload);

            Assert((string)json["model"] == "test-model", "Model mismatch");
            Assert((double)json["temperature"] == 0.5, "Temperature mismatch");
            Assert((int)json["max_tokens"] == 1000, "MaxTokens mismatch");
            Assert((bool)json["stream"] == false, "Stream mismatch");
            Assert(json["messages"] != null && ((JArray)json["messages"]).Count == 1, "Messages count mismatch");
            Assert(json["response_format"] == null, "Response format should be null when not specified");
            Assert(json["tools"] == null, "Tools should be null when not specified");
        }

        private static void TestOpenAICompatibleBuildPayloadJsonObject()
        {
            var req = new AIRequest
            {
                Model = "test-model",
                Messages = new List<ChatMessage> { new ChatMessage("user", "Give JSON") },
                ResponseFormat = "json_object"
            };

            var payload = OpenAICompatibleClient.BuildPayload(req, stream: true);
            var json = JObject.FromObject(payload);

            Assert((bool)json["stream"] == true, "Stream should be true");
            Assert(json["response_format"] != null, "response_format should not be null");
            Assert((string)json["response_format"]["type"] == "json_object", "response_format type should be json_object");
        }

        private static void TestOpenAICompatibleBuildPayloadToolsAndChoice()
        {
            var tools = new List<object>
            {
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_stock_quote",
                        description = "Get real-time stock price"
                    }
                }
            };

            var req = new AIRequest
            {
                Model = "tool-model",
                Messages = new List<ChatMessage> { new ChatMessage("user", "What is MSFT price?") },
                Tools = tools,
                ToolChoice = "auto"
            };

            var payload = OpenAICompatibleClient.BuildPayload(req, stream: false);
            var json = JObject.FromObject(payload);

            Assert(json["tools"] != null, "Tools should not be null");
            Assert(((JArray)json["tools"]).Count == 1, "Tools count should be 1");
            Assert((string)json["tool_choice"] == "auto", "ToolChoice should be auto");
        }

        private static void TestOpenAICompatibleBuildPayloadExtraParameters()
        {
            var req = new AIRequest
            {
                Model = "custom-model",
                Messages = new List<ChatMessage> { new ChatMessage("user", "Test extras") }
            };
            req.ExtraParameters["top_p"] = 0.9;
            req.ExtraParameters["seed"] = 42;

            var payload = OpenAICompatibleClient.BuildPayload(req, stream: false);
            var json = JObject.FromObject(payload);

            Assert((double)json["top_p"] == 0.9, "Extra parameter top_p missing or wrong");
            Assert((int)json["seed"] == 42, "Extra parameter seed missing or wrong");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception(message);
            }
        }
    }
}
