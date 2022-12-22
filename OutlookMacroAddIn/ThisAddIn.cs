using Outlook = Microsoft.Office.Interop.Outlook;


namespace OutlookMacroAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Примечание. Outlook больше не выдает это событие. Если имеется код, который 
            //    должно выполняться при завершении работы Outlook, см. статью на странице https://go.microsoft.com/fwlink/?LinkId=506785
        }

        public Outlook.Application GetApplication()
        { 
            return Application;
        }        

        public Outlook.Inspector GetInspector()
        {
            return GetApplication().ActiveInspector();
        }

        public Outlook.Explorer GetExplorer()
        {
            return GetApplication().ActiveExplorer();
        }       


        //public Outlook.MAPIFolder GetMAPIFolder()
        //{

        //    return (Outlook.MAPIFolder)Application.MAPIFolder;
        //}

        //public Outlook.MailItem GetMailItem()
        //{
        //    return (Outlook.MailItem)Application.MailItem;
        //}

        //public Outlook.TaskItem GetTaskItem()
        //{
        //    return (Outlook.TaskItem)Application.TaskItem;
        //}

        //public Outlook.ContactItem GetContactItem()
        //{
        //    return (Outlook.ContactItem)Application.ContactItem;
        //}



        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
