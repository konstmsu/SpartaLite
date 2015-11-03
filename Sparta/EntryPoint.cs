using Microsoft.Office.Interop.Excel;
using Sparta.Controls;
using Sparta.Sheets;

namespace Sparta
{
    public partial class EntryPoint
    {
        private void OnStartup(object sender, System.EventArgs e)
        {
            var factory = new SheetFactory(Application);
            factory.ShowSheet(s => new ContentsSheet(s, factory));

            Visible = XlSheetVisibility.xlSheetVeryHidden;
        }

        private void OnShutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(OnStartup);
            this.Shutdown += new System.EventHandler(OnShutdown);
        }

        #endregion

    }
}
