﻿namespace GoContactSyncMod
{
    partial class ConfigurationManagerForm
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConfigurationManagerForm));
			this.btAdd = new System.Windows.Forms.Button();
			this.btEdit = new System.Windows.Forms.Button();
			this.btDel = new System.Windows.Forms.Button();
			this.btClose = new System.Windows.Forms.Button();
			this.lbProfiles = new System.Windows.Forms.CheckedListBox();
			this.SuspendLayout();
			// 
			// btAdd
			// 
			this.btAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btAdd.Location = new System.Drawing.Point(265, 12);
			this.btAdd.Name = "btAdd";
			this.btAdd.Size = new System.Drawing.Size(89, 23);
			this.btAdd.TabIndex = 1;
			this.btAdd.Text = "&Add";
			this.btAdd.UseVisualStyleBackColor = true;
			this.btAdd.Click += new System.EventHandler(this.btAdd_Click);
			// 
			// btEdit
			// 
			this.btEdit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btEdit.Location = new System.Drawing.Point(265, 41);
			this.btEdit.Name = "btEdit";
			this.btEdit.Size = new System.Drawing.Size(89, 23);
			this.btEdit.TabIndex = 2;
			this.btEdit.Text = "&Edit";
			this.btEdit.UseVisualStyleBackColor = true;
			this.btEdit.Click += new System.EventHandler(this.btEdit_Click);
			// 
			// btDel
			// 
			this.btDel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btDel.Location = new System.Drawing.Point(265, 85);
			this.btDel.Name = "btDel";
			this.btDel.Size = new System.Drawing.Size(89, 23);
			this.btDel.TabIndex = 3;
			this.btDel.Text = "&Delete";
			this.btDel.UseVisualStyleBackColor = true;
			this.btDel.Click += new System.EventHandler(this.btDel_Click);
			// 
			// btClose
			// 
			this.btClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btClose.Location = new System.Drawing.Point(265, 187);
			this.btClose.Name = "btClose";
			this.btClose.Size = new System.Drawing.Size(89, 23);
			this.btClose.TabIndex = 4;
			this.btClose.Text = "Close";
			this.btClose.UseVisualStyleBackColor = true;
			this.btClose.Click += new System.EventHandler(this.btClose_Click);
			// 
			// lbProfiles
			// 
			this.lbProfiles.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.lbProfiles.FormattingEnabled = true;
			this.lbProfiles.Location = new System.Drawing.Point(12, 12);
			this.lbProfiles.Name = "lbProfiles";
			this.lbProfiles.Size = new System.Drawing.Size(247, 169);
			this.lbProfiles.TabIndex = 0;
			// 
			// ConfigurationManagerForm
			// 
			this.AcceptButton = this.btClose;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.btClose;
			this.ClientSize = new System.Drawing.Size(366, 222);
			this.Controls.Add(this.lbProfiles);
			this.Controls.Add(this.btClose);
			this.Controls.Add(this.btDel);
			this.Controls.Add(this.btEdit);
			this.Controls.Add(this.btAdd);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.MinimumSize = new System.Drawing.Size(308, 206);
			this.Name = "ConfigurationManagerForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Configuration Manager";
			this.Load += new System.EventHandler(this.ConfigurationManagerForm_Load);
			this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btAdd;
        private System.Windows.Forms.Button btEdit;
        private System.Windows.Forms.Button btDel;
        private System.Windows.Forms.Button btClose;
        private System.Windows.Forms.CheckedListBox lbProfiles;
    }
}