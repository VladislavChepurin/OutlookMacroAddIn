using OutlookMacroAddIn.Serializable.Interfaces;

namespace OutlookMacroAddIn.Serializable.Entity
{
    public class ConvertToProjectSettings: IConvertToProjectSettings
    {
        public string FolderCreateProgect { get; set; }
        public ConvertToProjectSettings(string folderCreateProgect)
        {
            FolderCreateProgect = folderCreateProgect;
        }
    }
}
