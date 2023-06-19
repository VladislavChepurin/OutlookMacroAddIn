using Microsoft.Office.Interop.Outlook;
using OutlookMacroAddIn.Functions.Models;
using OutlookMacroAddIn.Serializable.Entity;
using OutlookMacroAddIn.Serializable.Interfaces;
using System;
using System.IO;

namespace OutlookMacroAddIn.Functions
{
    internal class ConvertToCalc : AbstractFunctions
    {
        private readonly IConvertToProjectSettings settings;
        public ConvertToCalc(AppSettings settings)
        {
            this.settings = settings;
        }

        public override void Start()
        {
            string folder;
            if (string.IsNullOrEmpty(settings.FolderCreateCalc))
            {
                folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            else
            {
                folder = settings.FolderCreateCalc;
            }

            if (Inspector == null || Inspector.CurrentItem == null)
                return;

            var mail = Inspector.CurrentItem;
            var subject = mail.Subject();
            var foldersModel = new FoldersModels() { RootFolders = Path.Combine(folder, subject) };

            if (mail.attachments.count > 0)
            {
                CreateDirectory(foldersModel);
                for (int i = 1; i <= mail.attachments.count; i++)
                {
                    mail.attachments[i].saveasfile
                        (Path.Combine(foldersModel.RootFolders, mail.attachments[i].filename));
                }
            }
            else
            {
                MessageInformation("В данном письме нет вложений, создание папки расчета невозможно!", "Нет вложений");
            }

        }
        private static void CreateDirectory(FoldersModels foldersModel)
        {
            var directory = new DirectoryInfo(foldersModel.RootFolders);

            if (!directory.Exists)
            {
                directory.Create();
            }
        }
    }
}
