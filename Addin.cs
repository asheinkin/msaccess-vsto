using MSScriptControl;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyAddin
{

    public partial class Addin : Form
    {
        dynamic app;
        MSScriptControl.ScriptControl scr;
        bool forceClose=false;
        int timeOut=0;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.MenuItem menuItem1;
        private System.Windows.Forms.MenuItem menuItem2;
        private System.ComponentModel.IContainer addinComponents;
        public Addin(dynamic app)
        {
            InitializeComponent();

            this.app = app;

            this.components = new System.ComponentModel.Container();
            this.contextMenu1 = new System.Windows.Forms.ContextMenu();
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.menuItem2 = new System.Windows.Forms.MenuItem();

            this.contextMenu1.MenuItems.AddRange(
                   new System.Windows.Forms.MenuItem[] { this.menuItem1 , this.menuItem2 });
            this.menuItem1.Index = 0;
            this.menuItem1.Text = "&Show";
            this.menuItem1.Click += MenuItem1_Click;
            this.menuItem2.Index = 1;
            this.menuItem2.Text = "E&xit";
            this.menuItem2.Click += MenuItem2_Click;

            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            notifyIcon1.Icon = this.Icon;
            notifyIcon1.ContextMenu = this.contextMenu1;
            notifyIcon1.Text = "Addin";
            notifyIcon1.Visible = true;
            notifyIcon1.Click += NotifyIcon1_Click;


            scr= new MSScriptControl.ScriptControl();
        }

        private void NotifyIcon1_Click(object sender, EventArgs e)
        {
            this.Show();
            this.Activate();
        }

        private void MenuItem2_Click(object sender, EventArgs e)
        {
            forceClose = true;
            this.Close();
            app.Quit(2);
        }

        private void MenuItem1_Click(object sender, EventArgs e)
        {         
            this.Show();
            this.Activate();
        }

        private void Addin_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();               
            }
        }



        private void Addin_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!forceClose) { 
                e.Cancel = true;
                this.Hide();
            }
        }

        private void Addin_Load(object sender, EventArgs e)
        {
           
        }

        private void Addin_Shown(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void setup()
        {
            scr.Language = "vbscript";
            scr.Reset();

            scr.UseSafeSubset = false;
            scr.SitehWnd = (int)this.Handle;
            (scr as IScriptControl).Timeout = this.timeOut; ;
            scr.AllowUI = true;
            scr.AddObject("Application", app, true);
            //scr.State = ScriptControlStates.Initialized;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            setup();
            try
            {
                scr.AddObject("wscript", new Ctx(scr,res) as dynamic, false);

                var src = this.src.Text;
                if (src[0] == '?')
                {
                    res.Text = scr.Eval( src.Substring(1)).ToString();
                }else
                {
                   scr.ExecuteStatement(src);
                   
                }


            } 
            catch(Exception exc)
            {
                res.Text = string.Format("{0}\r\n{1}",exc.Source,exc.Message);
            }
            finally
            {
                scr.Reset();
            }
        }
    }
}
