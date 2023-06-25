using OutlookMacroAddIn.Serializable.Interfaces;

namespace OutlookMacroAddIn.Serializable.Entity
{
    public class AppSettings: IAppSettings
    {
        public string FolderCreateProgect { get; set; }
        public string FolderCreateCalc { get; set; }

        public AppSettings(string folderCreateProgect, string folderCreateCalc)
        {
            FolderCreateProgect = folderCreateProgect;
            FolderCreateCalc = folderCreateCalc;
        }
    }
}
