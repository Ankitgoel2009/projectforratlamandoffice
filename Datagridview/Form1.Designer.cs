
namespace Datagridview
{
    partial class Form1
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
            this.textBox2 = new Datagridview.textBox();
            this.textBox1 = new Datagridview.textBox();
            this.custombutton2 = new Datagridview.custombutton();
            this.SuspendLayout();
            // 
            // custombutton1
            // 
            this.custombutton1.Location = new System.Drawing.Point(345, 61);
            this.custombutton1.Name = "custombutton1";
            this.custombutton1.Size = new System.Drawing.Size(321, 180);
            this.custombutton1.TabIndex = 2;
            this.custombutton1.Text = "Accounting Vouchers";
            this.custombutton1.UseVisualStyleBackColor = true;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(76, 89);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(100, 22);
            this.textBox2.TabIndex = 1;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(62, 43);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(100, 22);
            this.textBox1.TabIndex = 0;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // custombutton2
            // 
            this.custombutton2.Location = new System.Drawing.Point(389, 327);
            this.custombutton2.Name = "custombutton2";
            this.custombutton2.Size = new System.Drawing.Size(277, 55);
            this.custombutton2.TabIndex = 3;
            this.custombutton2.Text = "custombutton2";
            this.custombutton2.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.custombutton2);
            this.Controls.Add(this.custombutton1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private textBox textBox1;
        private textBox textBox2;
        private custombutton custombutton1;
        private custombutton custombutton2;
    }
}

