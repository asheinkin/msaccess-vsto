
namespace MyAddin
{
    partial class Addin
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Addin));
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel6 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.res = new System.Windows.Forms.TextBox();
            this.src = new System.Windows.Forms.TextBox();
            this.panel4.SuspendLayout();
            this.panel6.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1161, 10);
            this.panel1.TabIndex = 2;
            // 
            // panel3
            // 
            this.panel3.Location = new System.Drawing.Point(1194, 380);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(183, 100);
            this.panel3.TabIndex = 4;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.panel6);
            this.panel4.Controls.Add(this.src);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(0, 10);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1161, 626);
            this.panel4.TabIndex = 5;
            // 
            // panel6
            // 
            this.panel6.Controls.Add(this.button1);
            this.panel6.Controls.Add(this.res);
            this.panel6.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel6.Location = new System.Drawing.Point(0, 451);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(1161, 175);
            this.panel6.TabIndex = 4;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(813, 17);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(111, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Run (Cntrl-R)";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // res
            // 
            this.res.Dock = System.Windows.Forms.DockStyle.Left;
            this.res.Location = new System.Drawing.Point(0, 0);
            this.res.Multiline = true;
            this.res.Name = "res";
            this.res.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.res.Size = new System.Drawing.Size(795, 175);
            this.res.TabIndex = 0;
            // 
            // src
            // 
            this.src.AllowDrop = true;
            this.src.Dock = System.Windows.Forms.DockStyle.Fill;
            this.src.Location = new System.Drawing.Point(0, 0);
            this.src.Multiline = true;
            this.src.Name = "src";
            this.src.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.src.Size = new System.Drawing.Size(1161, 626);
            this.src.TabIndex = 2;
            this.src.DragDrop += new System.Windows.Forms.DragEventHandler(this.src_DragDrop);
            this.src.DragOver += new System.Windows.Forms.DragEventHandler(this.src_DragOver);
            this.src.KeyDown += new System.Windows.Forms.KeyEventHandler(this.src_KeyDown);
            // 
            // Addin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1161, 636);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Addin";
            this.Text = "Addin";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Addin_FormClosing);
            this.Shown += new System.EventHandler(this.Addin_Shown);
            this.Resize += new System.EventHandler(this.Addin_Resize);
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.TextBox src;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox res;
    }
}