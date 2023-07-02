using Microsoft.Office.Interop.Outlook;
using OutlookMacroAddIn.Functions.Models;
using OutlookMacroAddIn.Serializable.Entity;
using OutlookMacroAddIn.Serializable.Interfaces;
using OutlookMacroAddIn.Services;
using System;
using System.IO;

namespace OutlookMacroAddIn.Functions
{
    internal class ConvertToCalc : AbstractFunctions
    {
        //private readonly IAppSettings settings;
        //public ConvertToCalc(AppSettings settings)
        //{
        //    this.settings = settings;
        //}

        public override void Start()
        {
            //string folder;
            //if (string.IsNullOrEmpty(settings.FolderCreateCalc))
            //{
               var folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //}
            //else
            //{
            //    folder = settings.FolderCreateCalc;
            //}

            if (Inspector == null || Inspector.CurrentItem == null)
                return;

            var mail = Inspector.CurrentItem;
            var subject = mail.Subject();
            var currentFolder = Path.Combine(folder, subject);

            if (mail.attachments.count > 0)
            {
                CreateDirectory(currentFolder);
                for (int i = 1; i <= mail.attachments.count; i++)
                {
                    mail.attachments[i].saveasfile
                        (Path.Combine(currentFolder, mail.attachments[i].filename));
                }
                AutoClosingMessageBox.Show("Успешно создана папка расчета", "Готово", 4000);
            }
            else
            {
                MessageInformation("В данном письме нет вложений, создание папки расчета невозможно!", "Нет вложений");
            }

        }
        private static void CreateDirectory(string foldersModel)
        {
            var directory = new DirectoryInfo(foldersModel);

            if (!directory.Exists)
            {
                directory.Create();
            }
        }
    }
}
