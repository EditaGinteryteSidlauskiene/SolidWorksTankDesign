namespace SolidWorksTankDesign.MVP.Views
{
    public partial class NozzleWindowView
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
            this.BackArrowButton = new System.Windows.Forms.Button();
            this.ForwardArrowButton = new System.Windows.Forms.Button();
            this.CompartmentsConfigurationLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // BackArrowButton
            // 
            this.BackArrowButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.BackArrowButton.BackColor = System.Drawing.Color.Transparent;
            this.BackArrowButton.Cursor = System.Windows.Forms.Cursors.Default;
            this.BackArrowButton.FlatAppearance.BorderSize = 0;
            this.BackArrowButton.FlatAppearance.CheckedBackColor = System.Drawing.Color.Transparent;
            this.BackArrowButton.Image = global::SolidWorksTankDesign.Properties.Resources.left;
            this.BackArrowButton.Location = new System.Drawing.Point(59, 588);
            this.BackArrowButton.Name = "BackArrowButton";
            this.BackArrowButton.Size = new System.Drawing.Size(56, 36);
            this.BackArrowButton.TabIndex = 8;
            this.BackArrowButton.UseVisualStyleBackColor = false;
            this.BackArrowButton.Click += new System.EventHandler(this.BackArrowButton_Click_1);
            // 
            // ForwardArrowButton
            // 
            this.ForwardArrowButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.ForwardArrowButton.BackColor = System.Drawing.Color.Transparent;
            this.ForwardArrowButton.Cursor = System.Windows.Forms.Cursors.Default;
            this.ForwardArrowButton.FlatAppearance.BorderSize = 0;
            this.ForwardArrowButton.FlatAppearance.CheckedBackColor = System.Drawing.Color.Transparent;
            this.ForwardArrowButton.Image = global::SolidWorksTankDesign.Properties.Resources.right;
            this.ForwardArrowButton.Location = new System.Drawing.Point(342, 588);
            this.ForwardArrowButton.Name = "ForwardArrowButton";
            this.ForwardArrowButton.Size = new System.Drawing.Size(56, 36);
            this.ForwardArrowButton.TabIndex = 9;
            this.ForwardArrowButton.UseVisualStyleBackColor = false;
            this.ForwardArrowButton.Click += new System.EventHandler(this.ForwardArrowButton_Click);
            // 
            // CompartmentsConfigurationLabel
            // 
            this.CompartmentsConfigurationLabel.AutoSize = true;
            this.CompartmentsConfigurationLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CompartmentsConfigurationLabel.Location = new System.Drawing.Point(41, 9);
            this.CompartmentsConfigurationLabel.Name = "CompartmentsConfigurationLabel";
            this.CompartmentsConfigurationLabel.Size = new System.Drawing.Size(141, 15);
            this.CompartmentsConfigurationLabel.TabIndex = 10;
            this.CompartmentsConfigurationLabel.Text = "Nozzle Configuration";
            // 
            // NozzleWindowView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.CompartmentsConfigurationLabel);
            this.Controls.Add(this.ForwardArrowButton);
            this.Controls.Add(this.BackArrowButton);
            this.Name = "NozzleWindowView";
            this.Size = new System.Drawing.Size(459, 1032);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BackArrowButton;
        private System.Windows.Forms.Button ForwardArrowButton;
        private System.Windows.Forms.Label CompartmentsConfigurationLabel;
    }
}
