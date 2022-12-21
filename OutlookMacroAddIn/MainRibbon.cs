using Microsoft.Office.Tools.Ribbon;
using OutlookMacroAddIn.Functions;
using System.IO;
using System;
using OutlookMacroAddIn.Serializable;

namespace OutlookMacroAddIn
{
    public partial class MainRibbon
    {
        private readonly string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config/appSettings.json");

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            AppSettingsDeserialize app = new AppSettingsDeserialize(jsonFilePath);
            var settings = app.GetSettingsModels();
            var convertToProjectSettings = settings.ConvertToProjectSettings;


            button1.Click += (s, a) =>
            {
                var convetrToProject = new ConvertToProject(convertToProjectSettings);
                convetrToProject.Start();
            };
        }
    }
}
