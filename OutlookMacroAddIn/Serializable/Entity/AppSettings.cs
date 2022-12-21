using System;

namespace OutlookMacroAddIn.Serializable.Entity
{
    [Serializable]
    public class AppSettings
    {
        public ConvertToProjectSettings ConvertToProjectSettings { get; set; }

        public AppSettings(ConvertToProjectSettings convertToProjectSettings)
        {
            ConvertToProjectSettings = convertToProjectSettings;
        }
    }
}
