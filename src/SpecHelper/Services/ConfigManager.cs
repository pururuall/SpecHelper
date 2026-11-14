using System;
using System.IO;
using Newtonsoft.Json;
using SpecHelper.Models;

namespace SpecHelper.Services
{
    public static class ConfigManager
    {
        private static readonly string ConfigDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SpecHelper"
        );

        private static readonly string ConfigFilePath = Path.Combine(ConfigDirectory, "config.json");

        public static PluginConfig Load()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var json = File.ReadAllText(ConfigFilePath);
                    var config = JsonConvert.DeserializeObject<PluginConfig>(json);
                    if (config != null)
                    {
                        ValidateOrFillDefaults(config);
                        return config;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ConfigManager] Ошибка чтения конфигурации: {ex.Message}");
            }

            var defaultConfig = new PluginConfig();
            ValidateOrFillDefaults(defaultConfig);
            Save(defaultConfig);
            return defaultConfig;
        }

        public static void Save(PluginConfig config)
        {
            try
            {
                if (!Directory.Exists(ConfigDirectory))
                {
                    Directory.CreateDirectory(ConfigDirectory);
                }

                var json = JsonConvert.SerializeObject(config, Formatting.Indented);
                File.WriteAllText(ConfigFilePath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ConfigManager] Ошибка записи конфигурации: {ex.Message}");
            }
        }

        private static void ValidateOrFillDefaults(PluginConfig config)
        {
            // Список потенциальных каталогов для баз данных при первом запуске
            string[] searchFolders = new string[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SpecHelper", "databases"),
                Path.Combine(Path.GetDirectoryName(typeof(ConfigManager).Assembly.Location) ?? "", "databases"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SpecHelper", "databases"),
                @"D:\EXCEL mibil\databases"
            };

            foreach (var folder in searchFolders)
            {
                if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder)) continue;

                if (string.IsNullOrWhiteSpace(config.MaterialsPath) && File.Exists(Path.Combine(folder, "Материалы.xlsx")))
                {
                    config.MaterialsPath = Path.Combine(folder, "Материалы.xlsx");
                }
                if (string.IsNullOrWhiteSpace(config.HardwarePath) && File.Exists(Path.Combine(folder, "Метизы.xlsx")))
                {
                    config.HardwarePath = Path.Combine(folder, "Метизы.xlsx");
                }
                if (string.IsNullOrWhiteSpace(config.ComponentsPath) && File.Exists(Path.Combine(folder, "Комплектующие.xlsx")))
                {
                    config.ComponentsPath = Path.Combine(folder, "Комплектующие.xlsx");
                }
            }
        }
    }
}
