﻿using Microsoft.Office.Interop.Outlook;
using OutlookMacroAddIn.Serializable.Interfaces;
using System;
using System.IO;
using OutlookMacroAddIn.Functions.Models;
using OutlookMacroAddIn.Serializable.Entity;
using OutlookMacroAddIn.Services;

namespace OutlookMacroAddIn.Functions
{
    internal class ConvertToProject : AbstractFunctions
    {

        private readonly IAppSettings settings;

        public ConvertToProject(AppSettings settings)
        {
            this.settings = settings;
        }

        public override void Start()
        {
            string folder;
            if (string.IsNullOrEmpty(settings.FolderCreateProgect))
            {
                folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            else
            {
                folder = settings.FolderCreateProgect;
            }


            if (Inspector == null || Inspector.CurrentItem == null)
                return;

            var mail = Inspector.CurrentItem;
            var subject = mail.Subject();

            var trimSubject = SubjectReplace.ProgectReplace(subject);
            var foldersModel = new FoldersModels(Path.Combine(folder, trimSubject)) ;
           
            if (mail.attachments.count > 0)
            {
                CreateDirectory(foldersModel);

                for (int i = 1; i <= mail.attachments.count; i++)
                {
                    mail.attachments[i].saveasfile
                        (Path.Combine(foldersModel.RootFolders, foldersModel.SourceDocumentationInfo, mail.attachments[i].filename));
                }
                AutoClosingMessageBox.Show("Успешно создана папка проекта", "Готово", 4000);
            }
            else
            {
                MessageInformation("В данном письме нет вложений, создание папки проекта невозможно!", "Нет вложений");
            }
        }

        private static void CreateDirectory(FoldersModels foldersModel)
        {
            var directory = new DirectoryInfo(foldersModel.RootFolders);
            
            if (!directory.Exists)
            {
                directory.Create();
                directory.CreateSubdirectory(foldersModel.SourceDocumentation);
                directory.CreateSubdirectory(foldersModel.SourceDocumentationInfo);
                directory.CreateSubdirectory(foldersModel.SourceDocumentationPassports);
                directory.CreateSubdirectory(foldersModel.SourceDocumentationCertificates);
                directory.CreateSubdirectory(foldersModel.AssemblyDocumentation);
                directory.CreateSubdirectory(foldersModel.AssemblyDocumentationDrawing);
                directory.CreateSubdirectory(foldersModel.ExecutiveDocumentation);
                directory.CreateSubdirectory(foldersModel.Photographs);
                directory.CreateSubdirectory(foldersModel.Logistics);
                directory.CreateSubdirectory(foldersModel.Complaints);
            }
        }
    }
}
