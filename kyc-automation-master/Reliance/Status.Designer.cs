using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace Reliance
{
    partial class Status
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Status));
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.content = new System.Windows.Forms.Label();
            this.axAcroPDF2 = new AxAcroPDFLib.AxAcroPDF();
            this.axAcroPDF1 = new AxAcroPDFLib.AxAcroPDF();
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF1)).BeginInit();
            this.SuspendLayout();
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            // 
            // content
            // 
            this.content.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.content.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.content.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold);
            this.content.Location = new System.Drawing.Point(29, 74);
            this.content.Name = "content";
            this.content.Size = new System.Drawing.Size(537, 64);
            this.content.TabIndex = 1;
            this.content.Text = "Checking...";
            this.content.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.content.UseMnemonic = false;
            // 
            // axAcroPDF2
            // 
            this.axAcroPDF2.Enabled = true;
            this.axAcroPDF2.Location = new System.Drawing.Point(29, 170);
            this.axAcroPDF2.Name = "axAcroPDF2";
            this.axAcroPDF2.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axAcroPDF2.OcxState")));
            this.axAcroPDF2.Size = new System.Drawing.Size(288, 288);
            this.axAcroPDF2.TabIndex = 3;
            // 
            // axAcroPDF1
            // 
            this.axAcroPDF1.Enabled = true;
            this.axAcroPDF1.Location = new System.Drawing.Point(323, 170);
            this.axAcroPDF1.Name = "axAcroPDF1";
            this.axAcroPDF1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axAcroPDF1.OcxState")));
            this.axAcroPDF1.Size = new System.Drawing.Size(288, 288);
            this.axAcroPDF1.TabIndex = 3;
            // 
            // Status
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(642, 496);
            this.Controls.Add(this.axAcroPDF1);
            this.Controls.Add(this.axAcroPDF2);
            this.Controls.Add(this.content);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "Status";
            this.Padding = new System.Windows.Forms.Padding(26, 74, 26, 25);
            this.Text = "Verifying Details";
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.axAcroPDF1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private BackgroundWorker backgroundWorker1;
        private Label content;
        private AxAcroPDFLib.AxAcroPDF axAcroPDF2;
        private AxAcroPDFLib.AxAcroPDF axAcroPDF1;
    }
}