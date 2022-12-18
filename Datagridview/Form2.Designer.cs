
namespace Datagridview
{
    partial class Form2
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
            this.custombutton1 = new Datagridview.custombutton();
            this.customControl11 = new Datagridview.CustomControl1();
            this.SuspendLayout();
            // 
            // custombutton1
            // 
            this.custombutton1.AutoSize = true;
            this.custombutton1.Location = new System.Drawing.Point(348, 82);
            this.custombutton1.Name = "custombutton1";
            this.custombutton1.Size = new System.Drawing.Size(101, 17);
            this.custombutton1.TabIndex = 0;
            this.custombutton1.Text = "custombutton1";
            // 
            // customControl11
            // 
            this.customControl11.Location = new System.Drawing.Point(361, 141);
            this.customControl11.Name = "customControl11";
            this.customControl11.Size = new System.Drawing.Size(189, 70);
            this.customControl11.TabIndex = 1;
            this.customControl11.Text = "customControl11";
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.customControl11);
            this.Controls.Add(this.custombutton1);
            this.Name = "Form2";
            this.Text = "Form2";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private custombutton custombutton1;
        private CustomControl1 customControl11;
    }
}