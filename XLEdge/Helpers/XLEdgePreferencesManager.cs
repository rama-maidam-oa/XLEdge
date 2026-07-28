using System;
using System.IO;
using System.Text.Json;
using XLEdge.Models;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    public sealed class XLEdgePreferencesManager
    {
        private static readonly Lazy<XLEdgePreferencesManager> _instance =
            new Lazy<XLEdgePreferencesManager>(() => new XLEdgePreferencesManager());

        public static XLEdgePreferencesManager Instance => _instance.Value;

        private readonly object _syncRoot = new object();

        private XLEdgeUserPreferences _loadedPreferences;

        private XLEdgePreferencesManager()
        {
        }

        public XLEdgeUserPreferences Current
        {
            get
            {
                lock (_syncRoot)
                {
                    return (_loadedPreferences ?? GetDefaultPreferencesFromAppState()).Clone();
                }
            }
        }

        public void Initialize()
        {
            lock (_syncRoot)
            {
                var defaults = GetDefaultPreferencesFromAppState();
                string filePath = GetUserPreferenceFilePath();

                bool shouldSave;
                var loaded = ReadPreferencesFromFile(filePath, defaults, out shouldSave);

                _loadedPreferences = loaded.Clone();
                ApplyToAppState(_loadedPreferences);

                if (shouldSave)
                {
                    WritePreferencesToFile(_loadedPreferences);
                }
            }
        }

        public void ApplyRuntime(XLEdgeUserPreferences preferences)
        {
            if (preferences == null)
            {
                throw new ArgumentNullException(nameof(preferences));
            }

            lock (_syncRoot)
            {
                ApplyToAppState(preferences);
            }
        }

        public void Save(XLEdgeUserPreferences preferences)
        {
            if (preferences == null)
            {
                throw new ArgumentNullException(nameof(preferences));
            }

            lock (_syncRoot)
            {
                _loadedPreferences = preferences.Clone();
                WritePreferencesToFile(_loadedPreferences);
                ApplyToAppState(_loadedPreferences);
            }
        }

        public void ResetRuntimeFromSaved()
        {
            lock (_syncRoot)
            {
                var source = _loadedPreferences ?? GetDefaultPreferencesFromAppState();
                ApplyToAppState(source);
            }
        }

        public XLEdgeUserPreferences GetFromAppState()
        {
            lock (_syncRoot)
            {
                return new XLEdgeUserPreferences
                {
                    ParameterValues = XLEdgeAppState.Instance.ParamDataSameSheet,
                    ScheduledOutputs = XLEdgeAppState.Instance.SchOutputsToSameSheet,
                    RefreshSync = XLEdgeAppState.Instance.RefreshSync,
                    ChangeSheetName = XLEdgeAppState.Instance.AllowSheetNameChanges,
                    CalendarCtrlDisplay = XLEdgeAppState.Instance.ShowCalendarControl,
                    OverrideFormats = XLEdgeAppState.Instance.OverrideFormats
                };
            }
        }

        private static string GetUserPreferenceFilePath()
        {
            return Path.Combine(XLEdgeAppPaths.XLEdgeLogsFolder, "xledgeuserpreferences.json");
        }

        private static XLEdgeUserPreferences GetDefaultPreferencesFromAppState()
        {
            return new XLEdgeUserPreferences
            {
                ParameterValues = XLEdgeAppState.Instance.ParamDataSameSheet,
                ScheduledOutputs = XLEdgeAppState.Instance.SchOutputsToSameSheet,
                RefreshSync = XLEdgeAppState.Instance.RefreshSync,
                ChangeSheetName = XLEdgeAppState.Instance.AllowSheetNameChanges,
                CalendarCtrlDisplay = XLEdgeAppState.Instance.ShowCalendarControl,
                OverrideFormats = XLEdgeAppState.Instance.OverrideFormats
            };
        }

        private static XLEdgeUserPreferences ReadPreferencesFromFile(
            string userPreferenceFile,
            XLEdgeUserPreferences defaults,
            out bool shouldSave)
        {
            shouldSave = false;

            if (!File.Exists(userPreferenceFile))
            {
                shouldSave = true;
                return defaults.Clone();
            }

            try
            {
                string fileContents = File.ReadAllText(userPreferenceFile);
                if (string.IsNullOrWhiteSpace(fileContents))
                {
                    shouldSave = true;
                    return defaults.Clone();
                }

                using (JsonDocument document = JsonDocument.Parse(fileContents))
                {
                    JsonElement root = document.RootElement;
                    bool missingProperty = false;

                    var preferences = new XLEdgeUserPreferences
                    {
                        ParameterValues = ReadBooleanProperty(root, "parameterValues", defaults.ParameterValues, ref missingProperty),
                        ScheduledOutputs = ReadBooleanProperty(root, "scheduledOutputs", defaults.ScheduledOutputs, ref missingProperty),
                        RefreshSync = ReadBooleanProperty(root, "refreshSync", defaults.RefreshSync, ref missingProperty),
                        ChangeSheetName = ReadBooleanProperty(root, "changeSheetName", defaults.ChangeSheetName, ref missingProperty),
                        CalendarCtrlDisplay = ReadBooleanProperty(root, "calendarCtrlDisplay", defaults.CalendarCtrlDisplay, ref missingProperty),
                        OverrideFormats = ReadBooleanProperty(root, "overrideFormats", defaults.OverrideFormats, ref missingProperty)
                    };

                    shouldSave = missingProperty;
                    return preferences;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed reading preference file. Falling back to defaults.");
                shouldSave = true;
                return defaults.Clone();
            }
        }

        private static bool ReadBooleanProperty(
            JsonElement root,
            string propertyName,
            bool defaultValue,
            ref bool missingProperty)
        {
            if (root.ValueKind == JsonValueKind.Object &&
                root.TryGetProperty(propertyName, out JsonElement propertyValue) &&
                (propertyValue.ValueKind == JsonValueKind.True || propertyValue.ValueKind == JsonValueKind.False))
            {
                return propertyValue.GetBoolean();
            }

            missingProperty = true;
            return defaultValue;
        }

        private static void WritePreferencesToFile(XLEdgeUserPreferences userPreferences)
        {
            string userPreferenceFile = GetUserPreferenceFilePath();

            Directory.CreateDirectory(Path.GetDirectoryName(userPreferenceFile));

            //var options = new JsonSerializerOptions
            //{
            //    WriteIndented = true
            //};

            //string fileContents = JsonSerializer.Serialize(userPreferences, options);
            string fileContents = SerializationHelper.SerializeToJson(userPreferences);
            
            File.WriteAllText(userPreferenceFile, fileContents);
        }

        private static void ApplyToAppState(XLEdgeUserPreferences userPreferences)
        {
            var appState = XLEdgeAppState.Instance;
            appState.ParamDataSameSheet = userPreferences.ParameterValues;
            appState.SchOutputsToSameSheet = userPreferences.ScheduledOutputs;
            appState.RefreshSync = userPreferences.RefreshSync;
            appState.AllowSheetNameChanges = userPreferences.ChangeSheetName;
            appState.ShowCalendarControl = userPreferences.CalendarCtrlDisplay;
            appState.OverrideFormats = userPreferences.OverrideFormats;
        }
    }
}