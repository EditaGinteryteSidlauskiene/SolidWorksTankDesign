namespace SolidWorksTankDesign.Windows
{
    partial class SettingsWindow
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ProjectSettings_Label = new System.Windows.Forms.Label();
            this.ProjectsFolder_Label = new System.Windows.Forms.Label();
            this.ProjectsFolder_BrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.Close_Button = new System.Windows.Forms.Button();
            this.ProjectsFolderTextBox = new System.Windows.Forms.TextBox();
            this.BrowseButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ProjectSettings_Label
            // 
            this.ProjectSettings_Label.AutoSize = true;
            this.ProjectSettings_Label.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ProjectSettings_Label.Location = new System.Drawing.Point(141, 25);
            this.ProjectSettings_Label.Name = "ProjectSettings_Label";
            this.ProjectSettings_Label.Size = new System.Drawing.Size(118, 20);
            this.ProjectSettings_Label.TabIndex = 0;
            this.ProjectSettings_Label.Text = "Project settings";
            // 
            // ProjectsFolder_Label
            // 
            this.ProjectsFolder_Label.AutoSize = true;
            this.ProjectsFolder_Label.Location = new System.Drawing.Point(24, 76);
            this.ProjectsFolder_Label.Name = "ProjectsFolder_Label";
            this.ProjectsFolder_Label.Size = new System.Drawing.Size(74, 13);
            this.ProjectsFolder_Label.TabIndex = 1;
            this.ProjectsFolder_Label.Text = "Projects folder";
            // 
            // Close_Button
            // 
            this.Close_Button.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Close_Button.Location = new System.Drawing.Point(158, 230);
            this.Close_Button.Name = "Close_Button";
            this.Close_Button.Size = new System.Drawing.Size(84, 30);
            this.Close_Button.TabIndex = 3;
            this.Close_Button.Text = "&Close";
            this.Close_Button.UseVisualStyleBackColor = true;
            // 
            // ProjectsFolderTextBox
            // 
            this.ProjectsFolderTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ProjectsFolderTextBox.Location = new System.Drawing.Point(128, 71);
            this.ProjectsFolderTextBox.Name = "ProjectsFolderTextBox";
            this.ProjectsFolderTextBox.Size = new System.Drawing.Size(239, 20);
            this.ProjectsFolderTextBox.TabIndex = 4;
            // 
            // BrowseButton
            // 
            this.BrowseButton.Location = new System.Drawing.Point(373, 69);
            this.BrowseButton.Name = "BrowseButton";
            this.BrowseButton.Size = new System.Drawing.Size(28, 23);
            this.BrowseButton.TabIndex = 5;
            this.BrowseButton.Text = "...";
            this.BrowseButton.UseVisualStyleBackColor = true;
            // 
            // SettingsWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.BrowseButton);
            this.Controls.Add(this.ProjectsFolderTextBox);
            this.Controls.Add(this.Close_Button);
            this.Controls.Add(this.ProjectsFolder_Label);
            this.Controls.Add(this.ProjectSettings_Label);
            this.Name = "SettingsWindow";
            this.Size = new System.Drawing.Size(429, 1060);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label ProjectSettings_Label;
        private System.Windows.Forms.Label ProjectsFolder_Label;
        private System.Windows.Forms.FolderBrowserDialog ProjectsFolder_BrowserDialog;
        private System.Windows.Forms.Button Close_Button;
        private System.Windows.Forms.TextBox ProjectsFolderTextBox;
        private System.Windows.Forms.Button BrowseButton;
    }
}
