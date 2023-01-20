using MSScriptControl;
using System;
using System.Collections;
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
        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr handle);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private  static extern IntPtr SendMessage(IntPtr hWnd, uint wMsg,              UIntPtr wParam, IntPtr lParam);
         

        private const uint WM_SYSCOMMAND  =                 0x0112;
        private   UIntPtr SC_RESTORE = (UIntPtr)0xF120;

        private Access.Application app;
        public string[] args;
        public bool forceClose=false;
        private bool show;
        private int timeOut;
        private TextWriter wr;
        public  System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.MenuItem menuItem1;
        private System.Windows.Forms.MenuItem menuItem2;
        private System.Windows.Forms.MenuItem menuItem3;
        private System.Windows.Forms.MenuItem menuItem4;
        public Addin(Microsoft.Office.Interop.Access.Application app,
                bool show,
                bool showVbe,
                bool topmost,
                int timeOut,
                TextWriter wr,
                string [] args,
                IDictionary env)
        {
            InitializeComponent();

            this.TopMost =topmost;
            this.show = show;
            this.timeOut = timeOut;
            this.wr = wr;
            args =Environment.GetCommandLineArgs();
            
            this.KeyPreview = true;
            this.app = app;

            this.components = new System.ComponentModel.Container();
            this.contextMenu1 = new System.Windows.Forms.ContextMenu();
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.menuItem2 = new System.Windows.Forms.MenuItem();
            this.menuItem3 = new System.Windows.Forms.MenuItem();
            this.menuItem4 = new System.Windows.Forms.MenuItem();

            this.contextMenu1.MenuItems.AddRange(
                   new System.Windows.Forms.MenuItem[] { this.menuItem1 , this.menuItem2, this.menuItem3, this.menuItem4 });
            this.contextMenu1.Popup += ContextMenu1_Popup;
            this.menuItem1.Index = 0;
            this.menuItem1.Text = "Addi&n";
            this.menuItem1.Click += MenuItem1_Click;

            this.menuItem2.Index = 1;
            this.menuItem2.Text = "VB&E";
            this.menuItem2.Click += MenuItem2_Click;

            this.menuItem3.Index = 2;
            this.menuItem3.Text = "&Access";
            this.menuItem3.Click += MenuItem3_Click;

            this.menuItem4.Index = 3;
            this.menuItem4.Text = "E&xit";
            this.menuItem4.Click += MenuItem4_Click;

            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            notifyIcon1.Icon = this.Icon;
            notifyIcon1.ContextMenu = this.contextMenu1;
            notifyIcon1.Text = "Addin";
            notifyIcon1.Visible = true;
            notifyIcon1.Click += NotifyIcon1_Click;

            this.Show();
            app.VBE.MainWindow.Visible =showVbe;
        }

        private void ContextMenu1_Popup(object sender, EventArgs e)
        {
            this.menuItem3.Checked = ! IsIconic((IntPtr)app.hWndAccessApp());
            this.menuItem2.Checked = app.VBE.MainWindow.Visible;
            this.menuItem1.Checked = this.Visible;
        }

        private void NotifyIcon1_Click(object sender, EventArgs e)
        {
            this.Show();
            this.Activate();
        }               

        private void MenuItem1_Click(object sender, EventArgs e)
        {
            this.Show();
            this.Activate();
            if (IsIconic(this.Handle)) SendMessage(this.Handle, WM_SYSCOMMAND, SC_RESTORE, IntPtr.Zero);
            SetForegroundWindow(this.Handle);
        }

        private void showVBE()
        {
            app.VBE.MainWindow.Visible = true;
            IntPtr hWnd = (IntPtr)app.VBE.MainWindow.HWnd;
            if (IsIconic(hWnd)) SendMessage(hWnd, WM_SYSCOMMAND, SC_RESTORE, IntPtr.Zero);
            SetForegroundWindow(hWnd);
        }

        private void showAccess()
        {
            IntPtr hWnd = (IntPtr)app.hWndAccessApp();
            if (IsIconic(hWnd)) SendMessage(hWnd, WM_SYSCOMMAND, SC_RESTORE, IntPtr.Zero);
            SetForegroundWindow(hWnd);
        }
        private void MenuItem2_Click(object sender, EventArgs e)
        {
            showVBE();
        }
        private void MenuItem3_Click(object sender, EventArgs e)
        {
            showAccess();
        }

        private void MenuItem4_Click(object sender, EventArgs e)
        {
            forceClose = true;
            this.notifyIcon1.Visible = false;
            this.notifyIcon1 = null;
            this.Close();
            app.DoCmd.Quit(Access.AcQuitOption.acQuitSaveNone);
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
                scr.AddObject("wsc", new Ctx(scr, res,wr) as dynamic, false);


                var wShell= Activator.CreateInstance(Type.GetTypeFromProgID("WScript.Shell"));
                scr.AddObject("WshShell", wShell, false);
                
                var fso = Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.FileSystemObject"));
                scr.AddObject("fso", wShell, false);

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
                wr.Flush();
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

        private void bAccess_Click(object sender, EventArgs e)
        {
            showAccess();
        }

        private void bVbe_Click(object sender, EventArgs e)
        {
            showVBE();
        }

        private void bCloseApp_Click(object sender, EventArgs e)
        {
            app.Quit();
        }

        private void Addin_Load(object sender, EventArgs e)
        {
            var local =Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var path = Path.Combine(local, "msaccess-myaddin","src.vbs");
            if (File.Exists(path))
            {
                src.Text = File.ReadAllText(path);              
            }
            
        }

        private void Addin_FormClosed(object sender, FormClosedEventArgs e)
        {
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var path = Path.Combine(local, "msaccess-myaddin");
            if (!Directory.Exists(path)){
                System.IO.Directory.CreateDirectory(path);
            };
            
            File.WriteAllText(Path.Combine(path, "src.vbs"), src.Text, Encoding.Unicode);

        }
    }
}



// class: wndclass_desked_gsk
// class: OMain
// class: VbaWindow
// caption: Immediate

// SendMessage(hwnd, WM_SYSCOMMAND, SC_RESTORE, 0)
// SetForegroundWindow(hwnd)
// SetActiveWindow(hwnd)
// SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)