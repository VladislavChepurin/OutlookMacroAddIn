using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace OutlookMacroAddIn.Functions
{
    internal abstract class AbstractFunctions
    {
        protected readonly Microsoft.Office.Interop.Outlook.Application Application = Globals.ThisAddIn.GetApplication();
        protected readonly Explorer Explorer = Globals.ThisAddIn.GetExplorers();
        protected readonly Inspector Inspector = Globals.ThisAddIn.GetInspectors();



        public abstract void Start();

        protected internal void MessageInformation(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        protected internal void MessageWarning(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        protected internal void MessageError(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }
    }
}
