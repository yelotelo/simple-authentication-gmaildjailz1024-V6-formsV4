using System;
using System.Drawing;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;


namespace simple_authentication_gmaildjailz1024_V6_formsV4
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        private Label label1;
        public TextBox textBox1;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            //this.label1 = new Label();


            //this.textBox1 = new TextBox();

            this.SuspendLayout();

            // label1


            // Form1
            this.AutoScaleDimensions = new SizeF(8F, 20F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1870, 800);
            this.Controls.Add(this.label1);


            this.Controls.Add(this.textBox1);

            this.Name = "Mytools";
            this.Text = "Mytools";

            this.Load += new EventHandler(this.Form1_Load);

            this.ResumeLayout(false);
            this.PerformLayout();
        }




    }
}
