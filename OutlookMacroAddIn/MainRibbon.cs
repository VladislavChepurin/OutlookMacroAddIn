using Microsoft.Office.Tools.Ribbon;
using OutlookMacroAddIn.Functions;

namespace OutlookMacroAddIn
{
    public partial class MainRibbon
    {
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            button1.Click += (s, a) =>
            {
                var convetrToProject = new ConvertToProject();
                convetrToProject.Start();
            };
        }
    }
}
