using System;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace MistralOfficeAddin
{
    /// <summary>
    /// User settings stored in HKCU\Software\MistralAIOffice.
    /// The API key is encrypted with DPAPI (per-user, per-machine).
    /// </summary>
    public sealed class Settings
    {
        private const string RegKey = "Software\\MistralAIOffice";
        private static readonly byte[] Entropy = new byte[] { 0x4D, 0x49, 0x53, 0x54, 0x52, 0x41, 0x4C }; // "MISTRAL"

        public string ApiKey = "";
        public string Model = "mistral-small-latest";
        public string BaseUrl = "https://api.mistral.ai/v1";
        public int TimeoutSeconds = 60;
        public string SystemPrompt = "You are a helpful AI assistant embedded in Microsoft Office. Be concise, accurate and practical.";
        public bool IncludeContextByDefault = true;

        public static Settings Load()
        {
            Settings s = new Settings();
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegKey))
                {
                    if (k == null) return s;
                    s.ApiKey = DecryptString(k.GetValue("ApiKeyEnc") as string);
                    object v;
                    v = k.GetValue("Model"); if (v is string && ((string)v).Length > 0) s.Model = (string)v;
                    v = k.GetValue("BaseUrl"); if (v is string && ((string)v).Length > 0) s.BaseUrl = (string)v;
                    v = k.GetValue("SystemPrompt"); if (v is string) s.SystemPrompt = (string)v;
                    v = k.GetValue("TimeoutSeconds"); if (v is int) s.TimeoutSeconds = Math.Max(5, Math.Min(300, (int)v));
                    v = k.GetValue("IncludeContextByDefault"); if (v is int) s.IncludeContextByDefault = ((int)v) != 0;
                }
            }
            catch
            {
                // Corrupt settings must never crash the host app.
            }
            return s;
        }

        public void Save()
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(RegKey))
            {
                k.SetValue("ApiKeyEnc", EncryptString(ApiKey), RegistryValueKind.String);
                k.SetValue("Model", Model, RegistryValueKind.String);
                k.SetValue("BaseUrl", BaseUrl, RegistryValueKind.String);
                k.SetValue("SystemPrompt", SystemPrompt, RegistryValueKind.String);
                k.SetValue("TimeoutSeconds", TimeoutSeconds, RegistryValueKind.DWord);
                k.SetValue("IncludeContextByDefault", IncludeContextByDefault ? 1 : 0, RegistryValueKind.DWord);
            }
        }

        private static string EncryptString(string plain)
        {
            if (string.IsNullOrEmpty(plain)) return "";
            byte[] bytes = ProtectedData.Protect(Encoding.UTF8.GetBytes(plain), Entropy, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(bytes);
        }

        private static string DecryptString(string cipher)
        {
            if (string.IsNullOrEmpty(cipher)) return "";
            try
            {
                byte[] bytes = Convert.FromBase64String(cipher);
                byte[] plain = ProtectedData.Unprotect(bytes, Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(plain);
            }
            catch
            {
                return "";
            }
        }
    }
}
