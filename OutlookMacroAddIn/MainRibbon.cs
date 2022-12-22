using Microsoft.Office.Tools.Ribbon;
using OutlookMacroAddIn.Functions;
using System.IO;
using System;
using OutlookMacroAddIn.Serializable;
using System.Threading;
using OutlookMacroAddIn.Services;

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

            var getRate = new GetCurrencyTsb
            {
                CurrencyHandler = ShowCurrencyPrice
            };
            //В новом потоке запускаем метод получения данных от Центробанка
            new Thread(() =>
            {
                getRate.Start();
            }).Start();
        }

        private void ShowCurrencyPrice(double usdCurrency, double euroCurrency, double cnhCurrency)
        {
            label1.Label = "Доллар = " + usdCurrency;
            label2.Label = "ЕВРО     = " + euroCurrency;
            label3.Label = "Юань    = " + cnhCurrency;
        }
    }
}
