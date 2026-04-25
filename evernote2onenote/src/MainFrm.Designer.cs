// Evernote2Onenote - imports Evernote notes to Onenote
// Copyright (C) 2014 - Stefan Kueng

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

namespace Evernote2Onenote
{
    partial class MainFrm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainFrm));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.infoText1 = new System.Windows.Forms.Label();
            this.infoText2 = new System.Windows.Forms.Label();
            this.progressIndicator = new System.Windows.Forms.ProgressBar();
            this.homeLink = new System.Windows.Forms.LinkLabel();
            this.versionLabel = new System.Windows.Forms.Label();
            this.btnENEXImport = new System.Windows.Forms.Button();
            this.importDatePicker = new System.Windows.Forms.DateTimePicker();
            this.datelabel = new System.Windows.Forms.Label();
            this.modifiedDateCheckbox = new System.Windows.Forms.CheckBox();
            this.lblImportFolder = new System.Windows.Forms.Label();
            this.txtImportFolder = new System.Windows.Forms.TextBox();
            this.btnBrowseImportFolder = new System.Windows.Forms.Button();
            this.lblLogFolder = new System.Windows.Forms.Label();
            this.txtLogFolder = new System.Windows.Forms.TextBox();
            this.btnBrowseLogFolder = new System.Windows.Forms.Button();
            this.lblNotebookName = new System.Windows.Forms.Label();
            this.txtNotebookName = new System.Windows.Forms.TextBox();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.GripMargin = new System.Windows.Forms.Padding(2, 2, 0, 2);
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(960, 48);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(71, 44);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(184, 44);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.ExitToolStripMenuItem_Click);
            // 
            // infoText1
            // 
            this.infoText1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.infoText1.AutoSize = true;
            this.infoText1.Location = new System.Drawing.Point(18, 407);
            this.infoText1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.infoText1.Name = "infoText1";
            this.infoText1.Size = new System.Drawing.Size(0, 25);
            this.infoText1.TabIndex = 3;
            // 
            // infoText2
            // 
            this.infoText2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.infoText2.AutoSize = true;
            this.infoText2.Location = new System.Drawing.Point(18, 442);
            this.infoText2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.infoText2.Name = "infoText2";
            this.infoText2.Size = new System.Drawing.Size(18, 25);
            this.infoText2.TabIndex = 4;
            this.infoText2.Text = " ";
            // 
            // progressIndicator
            // 
            this.progressIndicator.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressIndicator.Location = new System.Drawing.Point(15, 476);
            this.progressIndicator.Margin = new System.Windows.Forms.Padding(6);
            this.progressIndicator.Name = "progressIndicator";
            this.progressIndicator.Size = new System.Drawing.Size(926, 26);
            this.progressIndicator.TabIndex = 5;
            // 
            // homeLink
            // 
            this.homeLink.AutoSize = true;
            this.homeLink.Location = new System.Drawing.Point(24, 40);
            this.homeLink.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.homeLink.Name = "homeLink";
            this.homeLink.Size = new System.Drawing.Size(116, 25);
            this.homeLink.TabIndex = 8;
            this.homeLink.TabStop = true;
            this.homeLink.Text = "Homepage";
            this.homeLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.homeLink_LinkClicked);
            // 
            // versionLabel
            // 
            this.versionLabel.AutoSize = true;
            this.versionLabel.Location = new System.Drawing.Point(154, 40);
            this.versionLabel.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.versionLabel.Name = "versionLabel";
            this.versionLabel.Size = new System.Drawing.Size(91, 25);
            this.versionLabel.TabIndex = 9;
            this.versionLabel.Text = "Version:";
            // 
            // btnENEXImport
            // 
            this.btnENEXImport.Location = new System.Drawing.Point(331, 339);
            this.btnENEXImport.Margin = new System.Windows.Forms.Padding(6);
            this.btnENEXImport.Name = "btnENEXImport";
            this.btnENEXImport.Size = new System.Drawing.Size(553, 44);
            this.btnENEXImport.TabIndex = 13;
            this.btnENEXImport.Text = "Import ENEX Folder";
            this.btnENEXImport.UseVisualStyleBackColor = true;
            this.btnENEXImport.Click += new System.EventHandler(this.btnENEXImport_Click);
            // 
            // importDatePicker
            // 
            this.importDatePicker.Location = new System.Drawing.Point(331, 283);
            this.importDatePicker.Margin = new System.Windows.Forms.Padding(6);
            this.importDatePicker.MinDate = new System.DateTime(1799, 1, 1, 0, 0, 0, 0);
            this.importDatePicker.Name = "importDatePicker";
            this.importDatePicker.Size = new System.Drawing.Size(553, 31);
            this.importDatePicker.TabIndex = 14;
            // 
            // datelabel
            // 
            this.datelabel.AutoSize = true;
            this.datelabel.Location = new System.Drawing.Point(18, 287);
            this.datelabel.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.datelabel.Name = "datelabel";
            this.datelabel.Size = new System.Drawing.Size(294, 25);
            this.datelabel.TabIndex = 15;
            this.datelabel.Text = "only import notes newer than:";
            // 
            // modifiedDateCheckbox
            // 
            this.modifiedDateCheckbox.AutoSize = true;
            this.modifiedDateCheckbox.Checked = true;
            this.modifiedDateCheckbox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.modifiedDateCheckbox.Location = new System.Drawing.Point(24, 339);
            this.modifiedDateCheckbox.Margin = new System.Windows.Forms.Padding(6);
            this.modifiedDateCheckbox.Name = "modifiedDateCheckbox";
            this.modifiedDateCheckbox.Size = new System.Drawing.Size(288, 54);
            this.modifiedDateCheckbox.TabIndex = 16;
            this.modifiedDateCheckbox.Text = "Use Modified-Date of\r\nnotes as Date in Onenote";
            this.modifiedDateCheckbox.UseVisualStyleBackColor = true;
            // 
            // lblImportFolder
            // 
            this.lblImportFolder.AutoSize = true;
            this.lblImportFolder.Location = new System.Drawing.Point(18, 70);
            this.lblImportFolder.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lblImportFolder.Name = "lblImportFolder";
            this.lblImportFolder.Size = new System.Drawing.Size(216, 25);
            this.lblImportFolder.TabIndex = 17;
            this.lblImportFolder.Text = "Folder to import from:";
            // 
            // txtImportFolder
            // 
            this.txtImportFolder.Location = new System.Drawing.Point(15, 100);
            this.txtImportFolder.Margin = new System.Windows.Forms.Padding(6);
            this.txtImportFolder.Name = "txtImportFolder";
            this.txtImportFolder.Size = new System.Drawing.Size(814, 31);
            this.txtImportFolder.TabIndex = 18;
            // 
            // btnBrowseImportFolder
            // 
            this.btnBrowseImportFolder.Location = new System.Drawing.Point(835, 100);
            this.btnBrowseImportFolder.Margin = new System.Windows.Forms.Padding(6);
            this.btnBrowseImportFolder.Name = "btnBrowseImportFolder";
            this.btnBrowseImportFolder.Size = new System.Drawing.Size(110, 44);
            this.btnBrowseImportFolder.TabIndex = 19;
            this.btnBrowseImportFolder.Text = "Browse...";
            this.btnBrowseImportFolder.UseVisualStyleBackColor = true;
            this.btnBrowseImportFolder.Click += new System.EventHandler(this.btnBrowseImportFolder_Click);
            // 
            // lblLogFolder
            // 
            this.lblLogFolder.AutoSize = true;
            this.lblLogFolder.Location = new System.Drawing.Point(18, 142);
            this.lblLogFolder.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lblLogFolder.Name = "lblLogFolder";
            this.lblLogFolder.Size = new System.Drawing.Size(297, 25);
            this.lblLogFolder.TabIndex = 20;
            this.lblLogFolder.Text = "Folder to use for manifest log:";
            // 
            // txtLogFolder
            // 
            this.txtLogFolder.Location = new System.Drawing.Point(15, 172);
            this.txtLogFolder.Margin = new System.Windows.Forms.Padding(6);
            this.txtLogFolder.Name = "txtLogFolder";
            this.txtLogFolder.Size = new System.Drawing.Size(814, 31);
            this.txtLogFolder.TabIndex = 21;
            // 
            // btnBrowseLogFolder
            // 
            this.btnBrowseLogFolder.Location = new System.Drawing.Point(835, 172);
            this.btnBrowseLogFolder.Margin = new System.Windows.Forms.Padding(6);
            this.btnBrowseLogFolder.Name = "btnBrowseLogFolder";
            this.btnBrowseLogFolder.Size = new System.Drawing.Size(110, 44);
            this.btnBrowseLogFolder.TabIndex = 22;
            this.btnBrowseLogFolder.Text = "Browse...";
            this.btnBrowseLogFolder.UseVisualStyleBackColor = true;
            this.btnBrowseLogFolder.Click += new System.EventHandler(this.btnBrowseLogFolder_Click);
            // 
            // lblNotebookName
            // 
            this.lblNotebookName.AutoSize = true;
            this.lblNotebookName.Location = new System.Drawing.Point(18, 214);
            this.lblNotebookName.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lblNotebookName.Name = "lblNotebookName";
            this.lblNotebookName.Size = new System.Drawing.Size(240, 25);
            this.lblNotebookName.TabIndex = 23;
            this.lblNotebookName.Text = "Notebook to import into:";
            this.lblNotebookName.Click += new System.EventHandler(this.lblNotebookName_Click);
            // 
            // txtNotebookName
            // 
            this.txtNotebookName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtNotebookName.Location = new System.Drawing.Point(15, 244);
            this.txtNotebookName.Margin = new System.Windows.Forms.Padding(6);
            this.txtNotebookName.Name = "txtNotebookName";
            this.txtNotebookName.Size = new System.Drawing.Size(929, 31);
            this.txtNotebookName.TabIndex = 24;
            // 
            // MainFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(960, 512);
            this.Controls.Add(this.txtNotebookName);
            this.Controls.Add(this.lblNotebookName);
            this.Controls.Add(this.btnBrowseLogFolder);
            this.Controls.Add(this.txtLogFolder);
            this.Controls.Add(this.lblLogFolder);
            this.Controls.Add(this.btnBrowseImportFolder);
            this.Controls.Add(this.txtImportFolder);
            this.Controls.Add(this.lblImportFolder);
            this.Controls.Add(this.modifiedDateCheckbox);
            this.Controls.Add(this.datelabel);
            this.Controls.Add(this.importDatePicker);
            this.Controls.Add(this.btnENEXImport);
            this.Controls.Add(this.versionLabel);
            this.Controls.Add(this.homeLink);
            this.Controls.Add(this.progressIndicator);
            this.Controls.Add(this.infoText2);
            this.Controls.Add(this.infoText1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(6);
            this.MaximumSize = new System.Drawing.Size(1375, 583);
            this.MinimumSize = new System.Drawing.Size(905, 512);
            this.Name = "MainFrm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Evernote2Onenote";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainFrm_FormClosing);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.Label infoText1;
        private System.Windows.Forms.Label infoText2;
        private System.Windows.Forms.ProgressBar progressIndicator;
        private System.Windows.Forms.LinkLabel homeLink;
        private System.Windows.Forms.Label versionLabel;
        private System.Windows.Forms.Button btnENEXImport;
        private System.Windows.Forms.DateTimePicker importDatePicker;
        private System.Windows.Forms.Label datelabel;
        private System.Windows.Forms.CheckBox modifiedDateCheckbox;
        private System.Windows.Forms.Label lblImportFolder;
        private System.Windows.Forms.TextBox txtImportFolder;
        private System.Windows.Forms.Button btnBrowseImportFolder;
        private System.Windows.Forms.Label lblLogFolder;
        private System.Windows.Forms.TextBox txtLogFolder;
        private System.Windows.Forms.Button btnBrowseLogFolder;
        private System.Windows.Forms.Label lblNotebookName;
        private System.Windows.Forms.TextBox txtNotebookName;
    }
}
