using System;
using System.IO;

namespace OutlookMacroAddIn.Functions.Models
{
    public class FoldersModels
    {
        public string RootFolders {get; set;}
        public string SourceDocumentation { get; } = "1. Исходная документация";
        public string SourceDocumentationInfo { get; } = Path.Combine("1. Исходная документация", $"Инфо {DateTime.Now.ToString("dd.MM.yyyy")}");
        public string SourceDocumentationPassports { get; } = Path.Combine("1. Исходная документация", "Паспорта");
        public string SourceDocumentationCertificates { get; } = Path.Combine("1. Исходная документация", "Сертификаты");
        public string AssemblyDocumentation { get; } = "2. Сборочная документация";
        public string AssemblyDocumentationDrawing { get; } = Path.Combine("2. Сборочная документация", "Чертеж (DWG+PDF)");
        public string ExecutiveDocumentation { get; } = "3. Исполнительная документация";
        public string Photographs { get; } = "4. Фото";
        public string Logistics { get; } = "5. Логистика (Отгрузочные+подписанный документ с объекта)";
        public string Complaints { get; } = "6. Рекламации";
    }
}
