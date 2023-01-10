﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Access = Microsoft.Office.Interop.Access;
using Office = Microsoft.Office.Core;

namespace MyAddin
{
    public partial class ThisAddIn
    {
        public static Microsoft.Office.Interop.Access.Application app;
        public static Addin addin;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var args = Environment.GetCommandLineArgs();
            app = this.Application;
            string name = app.Name;
            string version = app.Version;
           // app.VBE.MainWindow.Visible = true;
            addin = new Addin(app);
            addin.Show();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            addin.Close();
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
