using System;
using System.IO;

namespace OutlookMacroAddIn.Functions.Models
{
    public class FoldersModels
    {
        public string RootFolders {get; set;}
        public string SourceDocumentation { get; set;}
        public string SourceDocumentationInfo { get; set; }
        public string SourceDocumentationPassports { get; set; }
        public string SourceDocumentationCertificates { get; set; }
        public string AssemblyDocumentation { get; set; }
        public string AssemblyDocumentationDrawing { get; set; }
        public string ExecutiveDocumentation { get; set;} 
        public string Photographs { get; set; } 
        public string Logistics { get; set; } 
        public string Complaints { get; set; }

        public FoldersModels(string rootFolders)
        {
            RootFolders = rootFolders;
            SourceDocumentation = "1. Исходная документация";
            SourceDocumentationInfo = Path.Combine(SourceDocumentation, $"Инфо {DateTime.Now:dd.MM.yyyy}");
            SourceDocumentationPassports = Path.Combine(SourceDocumentation, "Паспорта");
            SourceDocumentationCertificates = Path.Combine(SourceDocumentation, "Сертификаты");
            AssemblyDocumentation = "2. Сборочная документация";
            AssemblyDocumentationDrawing = Path.Combine(AssemblyDocumentation, "Чертеж (DWG+PDF)");
            ExecutiveDocumentation = "3. Исполнительная документация";
            Photographs = "4. Фото";
            Logistics = "5. Логистика (Отгрузочные+подписанный документ с объекта)";
            Complaints = "6. Рекламации";
        }
    }
}
