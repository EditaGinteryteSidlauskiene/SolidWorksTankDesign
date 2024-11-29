using SolidWorksTankDesign.Windows;

namespace SolidWorksTankDesign
{
    partial class TaskpaneHostUI
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
        private void InitializeComponent(string settingsFilePath)
        {
            this.RecognizeButton = new System.Windows.Forms.Button();
            this.NewButton = new System.Windows.Forms.Button();
            this.SettingsButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // RecognizeButton
            // 
            this.RecognizeButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RecognizeButton.Location = new System.Drawing.Point(48, 53);
            this.RecognizeButton.Name = "RecognizeButton";
            this.RecognizeButton.Size = new System.Drawing.Size(98, 36);
            this.RecognizeButton.TabIndex = 0;
            this.RecognizeButton.Text = "Recognize";
            this.RecognizeButton.UseVisualStyleBackColor = true;
            this.RecognizeButton.Click += RecognizeButton_Click;
            // 
            // NewButton
            // 
            this.NewButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.NewButton.Location = new System.Drawing.Point(266, 53);
            this.NewButton.Name = "NewButton";
            this.NewButton.Size = new System.Drawing.Size(98, 36);
            this.NewButton.TabIndex = 1;
            this.NewButton.Text = "New";
            this.NewButton.UseVisualStyleBackColor = true;
            this.NewButton.Click += NewButton_Click;
            // 
            // SettingsButton
            // 
            this.SettingsButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SettingsButton.Location = new System.Drawing.Point(155, 119);
            this.SettingsButton.Name = "SettingsButton";
            this.SettingsButton.Size = new System.Drawing.Size(98, 36);
            this.SettingsButton.TabIndex = 2;
            this.SettingsButton.Text = "Settings";
            this.SettingsButton.UseVisualStyleBackColor = true;
            this.SettingsButton.Click += SettingsButton_Click;
            // 
            // TaskpaneHostUI
            // 
            this.Controls.Add(this.SettingsButton);
            this.Controls.Add(this.NewButton);
            this.Controls.Add(this.RecognizeButton);
            this.Name = "TaskpaneHostUI";
            this.Size = new System.Drawing.Size(476, 1045);
            this.ResumeLayout(false);

        }
        #endregion

        //private System.Windows.Forms.Button RecognizeButton;
        //private System.Windows.Forms.Button NewButton;
        //private System.Windows.Forms.Button SettingsButton;
    }
}
