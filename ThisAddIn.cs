using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Access = Microsoft.Office.Interop.Access;
using Office = Microsoft.Office.Core;

namespace MyAddin
{
    public partial class ThisAddIn
    {
        public static Microsoft.Office.Interop.Access.Application app;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            dynamic ext = (sender as Microsoft.Office.Tools.AddIn).Extension;
            app = ext.Application;
            //string name = app.Name;
            //string version = app.Version;          
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            app = null;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
