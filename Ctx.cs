using System;
using System.Collections;
using System.Collections.Generic;
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
        public Ctx(MSScriptControl.ScriptControl scr,TextBox res)
        {
            this.scr = scr;
            this.res = res;
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
        


    }

    
}

