using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;

namespace MistralOfficeAddin
{
    public class ChatMessage
    {
        public string Role;
        public string Content;

        public ChatMessage(string role, string content)
        {
            Role = role;
            Content = content;
        }
    }

    /// <summary>
    /// Minimal Mistral La Plateforme client (chat completions + model list).
    /// Synchronous by design — callers run it on a background thread and
    /// marshal results back to the Office UI thread.
    /// </summary>
    public static class MistralClient
    {
        static MistralClient()
        {
            // Office 2010-era machines may still run a .NET that defaults to
            // TLS 1.0; force TLS 1.2 (and 1.3 where available) for api.mistral.ai.
            // Note: (SecurityProtocolType)3072 is TLS 1.2.  The enum member does
            // not exist in .NET 4.0, but the runtime supports it if 4.5+ is
            // installed.  (SecurityProtocolType)12288 is TLS 1.3 (Win10 1903+).
            try { ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072; } catch { }
            try { ServicePointManager.SecurityProtocol |= (SecurityProtocolType)12288; } catch { }
        }

        /// <summary>
        /// Creates a JavaScriptSerializer with MaxJsonLength raised to avoid
        /// truncation on large chat sessions or document context.
        /// </summary>
        private static JavaScriptSerializer CreateSerializer()
        {
            JavaScriptSerializer ser = new JavaScriptSerializer();
            ser.MaxJsonLength = int.MaxValue;
            return ser;
        }

        /// <summary>
        /// Calls POST {baseUrl}/chat/completions. Returns the assistant message
        /// text, or null with 'error' populated on failure.
        /// </summary>
        public static string ChatCompletion(string baseUrl, string apiKey, string model,
            IList<ChatMessage> messages, int timeoutSeconds, out string error)
        {
            error = null;
            if (string.IsNullOrEmpty(apiKey))
            {
                error = "No API key configured. Open the Mistral AI tab, click Settings and paste your API key.";
                return null;
            }

            List<object> msgs = new List<object>();
            foreach (ChatMessage m in messages)
            {
                Dictionary<string, object> d = new Dictionary<string, object>();
                d["role"] = m.Role;
                d["content"] = m.Content;
                msgs.Add(d);
            }
            Dictionary<string, object> body = new Dictionary<string, object>();
            body["model"] = model;
            body["messages"] = msgs;
            body["temperature"] = 0.7;
            body["stream"] = false;

            string json = CreateSerializer().Serialize(body);
            string responseText = Request(baseUrl.TrimEnd('/') + "/chat/completions", "POST",
                apiKey, json, timeoutSeconds, out error);
            if (responseText == null) return null;

            try
            {
                Dictionary<string, object> root = CreateSerializer().Deserialize<Dictionary<string, object>>(responseText);

                // JavaScriptSerializer deserializes JSON arrays as ArrayList,
                // NOT as object[].  The original code cast to object[] which
                // would return null and misreport "no choices".
                ArrayList choicesAL = root.ContainsKey("choices") ? root["choices"] as ArrayList : null;
                if (choicesAL == null || choicesAL.Count == 0)
                {
                    error = "Unexpected response from Mistral (no choices): " + Truncate(responseText, 300);
                    return null;
                }
                Dictionary<string, object> choice = choicesAL[0] as Dictionary<string, object>;
                if (choice == null)
                {
                    error = "Unexpected response from Mistral (choice not an object): " + Truncate(responseText, 300);
                    return null;
                }
                Dictionary<string, object> message = choice.ContainsKey("message")
                    ? choice["message"] as Dictionary<string, object> : null;
                if (message == null)
                {
                    error = "Unexpected response from Mistral (no message in choice): " + Truncate(responseText, 300);
                    return null;
                }
                string content = message.ContainsKey("content") ? message["content"] as string : null;
                if (content == null)
                {
                    error = "Unexpected response from Mistral (no content): " + Truncate(responseText, 300);
                    return null;
                }
                return content;
            }
            catch (Exception ex)
            {
                error = "Could not parse Mistral response: " + ex.Message;
                return null;
            }
        }

        /// <summary>
        /// Calls GET {baseUrl}/models — used by the Settings dialog's
        /// "Test Connection" button. Returns a summary string on success,
        /// null with 'error' populated on failure.
        /// </summary>
        public static string TestConnection(string baseUrl, string apiKey, int timeoutSeconds, out string error)
        {
            error = null;
            if (string.IsNullOrEmpty(apiKey))
            {
                error = "Enter an API key first.";
                return null;
            }
            string responseText = Request(baseUrl.TrimEnd('/') + "/models", "GET",
                apiKey, null, timeoutSeconds, out error);
            if (responseText == null) return null;

            try
            {
                Dictionary<string, object> root = CreateSerializer().Deserialize<Dictionary<string, object>>(responseText);
                ArrayList data = root.ContainsKey("data") ? root["data"] as ArrayList : null;
                if (data == null) return "Connected. (Model list not parsable.)";
                List<string> ids = new List<string>();
                foreach (object o in data)
                {
                    Dictionary<string, object> d = o as Dictionary<string, object>;
                    if (d != null && d.ContainsKey("id") && d["id"] is string) ids.Add((string)d["id"]);
                }
                return "Connected. " + ids.Count + " models available: " + string.Join(", ", ids.ToArray());
            }
            catch (Exception ex)
            {
                return "Connected (response parse note: " + ex.Message + ")";
            }
        }

        private static string Request(string url, string method, string apiKey,
            string jsonBody, int timeoutSeconds, out string error)
        {
            error = null;
            try
            {
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = method;
                req.ContentType = "application/json";
                req.Accept = "application/json";
                req.UserAgent = "MistralOfficeAddin/1.0";
                req.Timeout = Math.Max(5, timeoutSeconds) * 1000;
                req.ReadWriteTimeout = Math.Max(5, timeoutSeconds) * 1000;
                req.Headers.Add("Authorization", "Bearer " + apiKey);

                if (jsonBody != null)
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(jsonBody);
                    req.ContentLength = bytes.Length;
                    using (Stream s = req.GetRequestStream())
                    {
                        s.Write(bytes, 0, bytes.Length);
                    }
                }

                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                {
                    return sr.ReadToEnd();
                }
            }
            catch (WebException ex)
            {
                HttpWebResponse resp = ex.Response as HttpWebResponse;
                string body = "";
                if (ex.Response != null)
                {
                    try
                    {
                        using (StreamReader sr = new StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8))
                        {
                            body = sr.ReadToEnd();
                        }
                    }
                    catch { }
                }
                if (resp != null)
                {
                    int code = (int)resp.StatusCode;
                    if (code == 401)
                    {
                        error = "HTTP 401 Unauthorized \u2014 check your Mistral API key in Settings.";
                    }
                    else if (code == 429)
                    {
                        // Parse Retry-After header for better user guidance
                        string retryAfter = resp.Headers["Retry-After"];
                        string wait = "";
                        if (!string.IsNullOrEmpty(retryAfter))
                        {
                            wait = " Retry after " + retryAfter + " seconds.";
                        }
                        error = "HTTP 429 Rate limited \u2014 the Mistral free tier allows limited requests per second/minute." +
                            wait + " Wait a moment and try again." +
                            (body.Length > 0 ? " Details: " + Truncate(body, 200) : "");
                    }
                    else if (code == 404)
                    {
                        error = "HTTP 404 Not Found \u2014 check the Base URL and the model name in Settings." +
                            (body.Length > 0 ? " Details: " + Truncate(body, 200) : "");
                    }
                    else
                    {
                        error = "HTTP " + code + " from Mistral." +
                            (body.Length > 0 ? " Details: " + Truncate(body, 200) : "");
                    }
                }
                else
                {
                    error = "Network error: " + ex.Message;
                }
                return null;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return null;
            }
        }

        private static string Truncate(string s, int len)
        {
            if (s == null) return "";
            s = s.Replace("\r", " ").Replace("\n", " ").Trim();
            if (s.Length <= len) return s;
            return s.Substring(0, len) + "...";
        }
    }
}
