using MSScriptControl;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Access=Microsoft.Office.Interop.Access;
namespace MyAddin
{
    public partial class Addin : System.Windows.Forms.Form
    {
        private Access.Application app;
        public string[] args;
        private bool forceClose=false;
        private bool show;
        private int timeOut=60000;
        public  System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.MenuItem menuItem1;
        private System.Windows.Forms.MenuItem menuItem2;

        public Addin(Microsoft.Office.Interop.Access.Application app,bool show)
        {
            InitializeComponent();
            this.show = show;
            args=Environment.GetCommandLineArgs();
            
            this.KeyPreview = true;
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

            this.Show();              
        }

        private void NotifyIcon1_Click(object sender, EventArgs e)
        {
            this.Show();
            this.Activate();
        }

        private void MenuItem2_Click(object sender, EventArgs e)
        {
            forceClose = true;
            this.notifyIcon1.Visible = false;
            this.notifyIcon1 = null;
            this.Close();
            app.DoCmd.Quit(Access.AcQuitOption.acQuitSaveNone);
        }

        private void MenuItem1_Click(object sender, EventArgs e)
        {         
            this.Show();
            this.Activate();
        }
        
        private void Script_Timeout()
        {
//            throw new NotImplementedException();
        }

        private void Script_Error()
        {
 //           throw new NotImplementedException();
        }
        private void runClick(object sender, EventArgs e)
        {
            ScriptControl scr= new ScriptControl();
            try
            {
                ((DScriptControlSource_Event)scr).Error += new DScriptControlSource_ErrorEventHandler(Script_Error);
                ((DScriptControlSource_Event)scr).Timeout += new DScriptControlSource_TimeoutEventHandler(Script_Timeout);

                scr.Language = "vbscript";                
                scr.UseSafeSubset = false;
                scr.SitehWnd = (int)this.Handle;
                (scr as IScriptControl).Timeout = this.timeOut; ;
                scr.AllowUI = true;
                scr.AddObject("Application", app, true);
                scr.AddObject("wscript", new Ctx(scr,res) as dynamic, false);

                var src = this.src.Text;
                if (src[0] == '?')
                {
                    res.AppendText( scr.Eval( src.Substring(1)).ToString()+Environment.NewLine);
                }else
                {
                   scr.ExecuteStatement(src);                   
                }
            } 
            catch(Exception ex)
            {
                IScriptControl iscriptControl = scr as MSScriptControl.IScriptControl;
                if (ex.Message.StartsWith("QUIT: "))
                {                     
                    this.res.AppendText($"\r\n{ex.Message}\r\n  at line {iscriptControl.Error.Line}\r\n");
                }
                 else
                {                    
                    this.res.AppendText(
                        $"\r\nERROR : {ex.Message} {ex.HResult}"
                        + Environment.NewLine + "  Description  : " + iscriptControl.Error.Description
                        + Environment.NewLine + "  Number       : " + iscriptControl.Error.Number
                        + Environment.NewLine + "  Source       : " + iscriptControl.Error.Source
                        + Environment.NewLine + "  Line of error: " + iscriptControl.Error.Line
                        + Environment.NewLine + "  Col  of error: " + iscriptControl.Error.Column
                        + Environment.NewLine + "  Code error   : " + iscriptControl.Error.Text
                        + Environment.NewLine
                        + Environment.NewLine)
                        ;                   
                }
                iscriptControl.Error.Clear();
            }
            
            finally
            {
                scr.Reset();
                Marshal.ReleaseComObject(scr);
                scr = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void Addin_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!forceClose)
            {
                e.Cancel = true;
                this.Hide();
            } 
        }

        private void Addin_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
            }
        }

        private void Addin_Shown(object sender, EventArgs e)
        {
           if (!this.show)  this.Hide();
        }

        private void src_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;
            else if (e.Data.GetDataPresent(DataFormats.Text))
            {
                e.Effect = DragDropEffects.Copy;
                Point p = src.PointToClient(new Point(e.X, e.Y));
                int index = src.GetCharIndexFromPosition(p);
                Point cp = src.GetPositionFromCharIndex(index);
                char c = src.GetCharFromPosition(p);
                using (var g = src.CreateGraphics())
                {
                    var s = g.MeasureString(c.ToString(), src.Font);
                    if (p.X > cp.X + s.Width / 2) ++index;
                }
                src.SelectionStart = index;
                src.SelectionLength = 0;
                src.Focus();
            }
            else
                e.Effect = DragDropEffects.None;
        }

        private void src_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Effect == DragDropEffects.Copy)
            {
              src.Text =src.Text.Insert(src.SelectionStart, (string)e.Data.GetData(DataFormats.Text));
              //int index = src.GetCharIndexFromPosition(src.PointToClient(Cursor.Position));
            }
            else if (e.Effect == DragDropEffects.Link)
            {
                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[]; // get all files droppeds  
                if (files != null && files.Any())
                {
                    src.Text = File.ReadAllText(files.First());                    
                }
            }
        }

        private void Addin_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Modifiers == Keys.Control && e.Modifiers != Keys.Alt  && e.Modifiers !=Keys.Shift   )
            {
                if (e.KeyCode == Keys.R)
                {
                    runClick(sender, e);
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                }
                else if (e.KeyCode == Keys.E)
                {
                    this.res.Text = "";
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                }
            } 
        }

        private void clearClick(object sender, EventArgs e)
        {
            this.res.Text = "";
        }
    }
}
