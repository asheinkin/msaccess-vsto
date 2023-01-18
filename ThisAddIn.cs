﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using Access = Microsoft.Office.Interop.Access;
using Office = Microsoft.Office.Core;

namespace MyAddin
{
    public partial class ThisAddIn
    {
        [DllImport("kernel32.dll")]
        private static extern uint GetPrivateProfileSection(string lpAppName, byte[] lpszReturnBuffer, uint nSize, string lpFileName);

        public static Microsoft.Office.Interop.Access.Application app;
        public static Addin addin;
        TextWriter wr;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var args = Environment.GetCommandLineArgs();

            var iniPath=Path.Combine(Environment.CurrentDirectory, "myaddin.ini");
            var ini = GetKeys(iniPath, "app");
            var env= Environment.GetEnvironmentVariables();

            app = this.Application;
            string name = app.Name;
            string version = app.Version;
            int timeOut = 60000;
            if (ini.ContainsKey("timeout"))
            {
                timeOut= int.Parse(ini["timeout"]);
            }
            app.VBE.MainWindow.Visible = ini.ContainsKey("vbe") && getBool(ini["vbe"]);       
            if (ini.ContainsKey("output") && ini["output"]!="")
            {
                bool f;
                if (ini.ContainsKey("append")) {
                    f = getBool(ini["append"]);
                } else {
                    f = true;
                }
                wr =  new StreamWriter(ini["output"], f, Encoding.Default);
            } else
            {
                wr = new DummyWriter();
            }

            addin = new Addin(app, ini.ContainsKey("show") && getBool(ini["show"]), timeOut,wr,args,env);             
            addin.Show();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            wr.Close();
            app = null;
            if (addin.notifyIcon1 != null)
            {
                addin.notifyIcon1.Visible = false;
            }
            //Dispose();
            //addin.Close();           
        }
        private Dictionary<string,string> GetKeys(string iniFile, string category)
        {
            // string text = System.IO.File.ReadAllText(iniFile,Encoding.UTF8);
            const uint MAX_BUFFER = 32767;
            byte[] buffer = new byte[MAX_BUFFER];
            
            uint bytesReturned = GetPrivateProfileSection(category, buffer, MAX_BUFFER, iniFile);
            String[] tmp = Encoding.UTF8.GetString(buffer).Trim('\0').Split('\0');
            var result = new Dictionary<string, string>();
            foreach (String entry in tmp)
            {
                if (!string.IsNullOrWhiteSpace(entry))
                {
                    var splitted =entry.Split(new char[] { '='},2);
                    if (splitted.Length > 1)
                    {
                        var zz0 = Regex.Match(splitted[0], @"^\s*(.+)\s*$");
                        var zz1 = Regex.Match(splitted[1], @"^\s*(.*)\s*$");
                        if (result.ContainsKey(zz0.Value))
                        {
                            result[zz0.Value] = zz1.Value;
                        }
                        else
                        {
                            result.Add(zz0.Value, zz1.Value);
                        }
                    }
                }
            }

            return result;
        }

        private bool getBool(string val)
        {
            if (val =="0" 
                || string.Compare(val,"false", true)==0
                || string.Compare(val, "f", true) == 0
                || string.Compare(val, "n", true) == 0
                || string.Compare(val, "no", true) == 0)
            {
                return false;
            } else
            {
                return true;
            }
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
