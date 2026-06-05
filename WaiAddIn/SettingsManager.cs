using System;
using System.IO;
using Newtonsoft.Json;

namespace WaiAddIn
{
    public class SettingsData
    {
        public string ApiKey { get; set; } = "";
        public string BaseUrl { get; set; } = "https://api.deepseek.com/chat/completions";
        public string Model { get; set; } = "deepseek-v4-flash";
        public int MaxTokens { get; set; } = 2048;
        public double Temperature { get; set; } = 0.2;
        public bool EditModeDefault { get; set; } = true;
        public int SelectionBoxHeight { get; set; } = 96;
        public int PromptBoxHeight { get; set; } = 140;
        public int ResponseBoxHeight { get; set; } = 210;
        public bool EnableAutocomplete { get; set; } = false;
        public string LastSessionId { get; set; } = "";

        public void Save()
        {
            SettingsManager.Save(this);
        }
    }

    public class SettingsManager
    {
        private static readonly string SettingsPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "WAI", "settings.json");

        public static SettingsData Load()
        {
            try
            {
                var dir = Path.GetDirectoryName(SettingsPath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                if (File.Exists(SettingsPath))
                {
                    var json = File.ReadAllText(SettingsPath);
                    return JsonConvert.DeserializeObject<SettingsData>(json) ?? new SettingsData();
                }
            }
            catch { }
            return new SettingsData();
        }

        public static void Save(SettingsData settings)
        {
            try
            {
                var dir = Path.GetDirectoryName(SettingsPath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(SettingsPath, json);
            }
            catch { }
        }
    }
}
