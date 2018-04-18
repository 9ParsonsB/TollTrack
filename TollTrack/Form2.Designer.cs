namespace TollTrack
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
            this.txtout = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtout
            // 
            this.txtout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtout.Location = new System.Drawing.Point(0, 0);
            this.txtout.Multiline = true;
            this.txtout.Name = "txtout";
            this.txtout.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtout.Size = new System.Drawing.Size(563, 261);
            this.txtout.TabIndex = 0;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(563, 261);
            this.Controls.Add(this.txtout);
            this.Name = "Form2";
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtout;
    }
}