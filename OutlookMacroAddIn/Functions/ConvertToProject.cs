using System;
using System.IO;

namespace OutlookMacroAddIn.Functions
{
    internal class ConvertToProject : AbstractFunctions
    {
        //private readonly 
        public ConvertToProject()
        {

        }

        public override void Start()
        {
            throw new NotImplementedException();
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
