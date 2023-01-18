using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyAddin
{
    [ComVisible(true)]
    public class Ctx
    {
        MSScriptControl.ScriptControl scr;
        TextBox res;
        TextWriter wr;
        public Ctx(MSScriptControl.ScriptControl scr,TextBox res, TextWriter wr)
        {
            this.scr = scr;
            this.res = res;
            this.wr = wr;
        }
        public void echo(params object[]  msg)
        {
            foreach(object o in msg)
            {
                res.AppendText(o.ToString());
            }
            res.AppendText(Environment.NewLine);
        }
        public void sleep(int n)
        {
            Thread.Sleep(n);
        }
        public void quit(int rc=0)
        {
            throw new Exception("QUIT: "+rc);
        }
        public string command()
        {
            return Environment.CommandLine; 
        }

        public void output(params object[] msg)
        {
            foreach (object o in msg)
            {
                wr.Write(o.ToString());
            }
            
        }
        public void Write(params object[] msg)
        {
            foreach (object o in msg)
            {
                wr.Write(o.ToString());
            }
        }
        public void WriteLine(params object[] msg)
        {
            foreach (object o in msg)
            {
                wr.Write(o.ToString());
            }
            wr.WriteLine();
        }

        public void WriteBlankLines(int n)
        {
            for (int i= 0;i< n;i++)
            {
                            wr.WriteLine();
            }
        }

    }

    
}

