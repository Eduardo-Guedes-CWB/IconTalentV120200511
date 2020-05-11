namespace IconTalentV1
{
    partial class wfMain
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
            this.pnlTop1 = new System.Windows.Forms.Panel();
            this.btnNewDoc = new System.Windows.Forms.Button();
            this.tcMain = new System.Windows.Forms.TabControl();
            this.pnlTop1.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlTop1
            // 
            this.pnlTop1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlTop1.Controls.Add(this.btnNewDoc);
            this.pnlTop1.Location = new System.Drawing.Point(0, 2);
            this.pnlTop1.Name = "pnlTop1";
            this.pnlTop1.Size = new System.Drawing.Size(799, 32);
            this.pnlTop1.TabIndex = 3;
            // 
            // btnNewDoc
            // 
            this.btnNewDoc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNewDoc.Location = new System.Drawing.Point(3, 3);
            this.btnNewDoc.Name = "btnNewDoc";
            this.btnNewDoc.Size = new System.Drawing.Size(85, 19);
            this.btnNewDoc.TabIndex = 0;
            this.btnNewDoc.Text = "Nova Lista";
            this.btnNewDoc.UseVisualStyleBackColor = true;
            this.btnNewDoc.Click += new System.EventHandler(this.btnNewDoc_Click);
            // 
            // tcMain
            // 
            this.tcMain.Location = new System.Drawing.Point(4, 40);
            this.tcMain.Name = "tcMain";
            this.tcMain.SelectedIndex = 0;
            this.tcMain.ShowToolTips = true;
            this.tcMain.Size = new System.Drawing.Size(795, 413);
            this.tcMain.TabIndex = 4;
            // 
            // wfMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.tcMain);
            this.Controls.Add(this.pnlTop1);
            this.Name = "wfMain";
            this.Text = "Icon Talent";
            this.pnlTop1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel pnlTop1;
        private System.Windows.Forms.Button btnNewDoc;
        private System.Windows.Forms.TabControl tcMain;
    }
}

