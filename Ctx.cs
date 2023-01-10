using System;
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
                res.Text+= o.ToString();
            }
            res.Text += Environment.NewLine;
        }
        public void Sleep(int n)
        {
            Thread.Sleep(n);
        }
        public void quit(int rc=0)
        {
            throw new Exception("break");

        }
       
    }
}

