using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using SpecHelper.Models;

namespace SpecHelper.Services
{
    public class DatabaseWatcher : IDisposable
    {
        public event Action<CatalogType, string> FileChanged;

        private readonly Dictionary<CatalogType, FileSystemWatcher> _watchers = new Dictionary<CatalogType, FileSystemWatcher>();
        private readonly Dictionary<CatalogType, Timer> _debounceTimers = new Dictionary<CatalogType, Timer>();
        private readonly object _lock = new object();
        private int _debounceMs = 800;

        public DatabaseWatcher(int debounceMs = 800)
        {
            _debounceMs = Math.Max(200, debounceMs);
        }

        public void UpdatePaths(PluginConfig config)
        {
            lock (_lock)
            {
                _debounceMs = Math.Max(200, config.SyncDebounceMs);

                SetupWatcher(CatalogType.Materials, config.MaterialsPath, config.AutoSyncEnabled);
                SetupWatcher(CatalogType.Hardware, config.HardwarePath, config.AutoSyncEnabled);
                SetupWatcher(CatalogType.Components, config.ComponentsPath, config.AutoSyncEnabled);
            }
        }

        private void SetupWatcher(CatalogType type, string filePath, bool enabled)
        {
            if (_watchers.TryGetValue(type, out var existingWatcher))
            {
                existingWatcher.EnableRaisingEvents = false;
                existingWatcher.Dispose();
                _watchers.Remove(type);
            }

            if (!enabled || string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }

            try
            {
                var dir = Path.GetDirectoryName(filePath);
                var filename = Path.GetFileName(filePath);

                if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
                {
                    return;
                }

                var watcher = new FileSystemWatcher(dir, filename)
                {
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName
                };

                watcher.Changed += (s, e) => OnFileEvent(type, filePath);
                watcher.Created += (s, e) => OnFileEvent(type, filePath);
                watcher.Renamed += (s, e) => OnFileEvent(type, filePath);

                watcher.EnableRaisingEvents = true;
                _watchers[type] = watcher;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DatabaseWatcher] Не удалось создать FileSystemWatcher для {type}: {ex.Message}");
            }
        }

        private void OnFileEvent(CatalogType type, string filePath)
        {
            lock (_lock)
            {
                if (_debounceTimers.TryGetValue(type, out var timer))
                {
                    timer.Change(_debounceMs, Timeout.Infinite);
                }
                else
                {
                    var newTimer = new Timer((state) =>
                    {
                        var tuple = ((CatalogType Type, string Path))state;
                        FileChanged?.Invoke(tuple.Type, tuple.Path);
                    }, (type, filePath), _debounceMs, Timeout.Infinite);

                    _debounceTimers[type] = newTimer;
                }
            }
        }

        public void Dispose()
        {
            lock (_lock)
            {
                foreach (var watcher in _watchers.Values)
                {
                    try { watcher.EnableRaisingEvents = false; watcher.Dispose(); } catch { }
                }
                _watchers.Clear();

                foreach (var timer in _debounceTimers.Values)
                {
                    try { timer.Dispose(); } catch { }
                }
                _debounceTimers.Clear();
            }
        }
    }
}
