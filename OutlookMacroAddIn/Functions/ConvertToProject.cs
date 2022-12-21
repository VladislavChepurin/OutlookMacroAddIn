using OutlookMacroAddIn.Serializable.Interfaces;
using System;
using System.IO;

namespace OutlookMacroAddIn.Functions
{
    internal class ConvertToProject : AbstractFunctions
    {
               
        private readonly IConvertToProjectSettings settings;

        public ConvertToProject(IConvertToProjectSettings settings)
        {
            this.settings= settings;
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

            var inspector = Inspector.CurrentItem;
            string subject = inspector.Subject();

            CreateDirectory(folder, subject);

        }

        private void CreateDirectory(string rootDirectory, string path)
        {
            var directory = new DirectoryInfo(Path.Combine(rootDirectory, path));
            var dateTime = DateTime.Now;

            if (!directory.Exists)
            {
                directory.Create();
                directory.CreateSubdirectory("1. Исходная документация");

                directory.CreateSubdirectory(Path.Combine("1. Исходная документация", $"Инфо {dateTime.ToString("dd.MM.yyyy")}"));
                directory.CreateSubdirectory(Path.Combine("1. Исходная документация", "Паспорта"));
                directory.CreateSubdirectory(Path.Combine("1. Исходная документация", "Сертификаты_"));

                directory.CreateSubdirectory("2. Сборочная документация");
                directory.CreateSubdirectory(Path.Combine("2. Сборочная документация", "Чертеж (DWG+PDF)"));

                directory.CreateSubdirectory("3. Исполнительная документация");
                directory.CreateSubdirectory("4. Фото");
                directory.CreateSubdirectory("5. Логистика (Отгрузочные+подписанный документ с объекта)");
                directory.CreateSubdirectory("6. Рекламации");
            }
        }
    }
}
