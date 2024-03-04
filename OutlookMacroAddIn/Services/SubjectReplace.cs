using System;

namespace OutlookMacroAddIn.Services
{
    internal static class SubjectReplace
    {   
        public static string ProgectReplace(string subject)
        {     
            return CalcReplace(
                subject.Replace("НОВЫЙ ПРОЕКТ ", String.Empty)                                        
                .Replace("Re: ", String.Empty)                                       
                .Replace("Fw: ", String.Empty)                                       
                .Replace("Fwd: ", String.Empty));
        }

        public static string CalcReplace(string subject)
        {
            return subject
                .Replace("<", "_")
                .Replace(">", "_")
                .Replace(":", "_")
                .Replace("\"", "_")
                .Replace("/", "_")
                .Replace("\\", "_")
                .Replace("|", "_")
                .Replace("?", "_")
                .Replace("*", "_");
        }
    }
}
