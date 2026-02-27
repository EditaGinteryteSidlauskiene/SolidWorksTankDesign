namespace SolidWorksTankDesign.Windows
{
    partial class CompartmentWindowView
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
            this.NewCompartmentButton = new System.Windows.Forms.Button();
            this.BackArrowButton = new System.Windows.Forms.Button();
            this.ForwardArrowButton = new System.Windows.Forms.Button();
            this.CompartmentsConfigurationLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // NewCompartmentButton
            // 
            this.NewCompartmentButton.BackColor = System.Drawing.Color.Transparent;
            this.NewCompartmentButton.FlatAppearance.BorderSize = 0;
            this.NewCompartmentButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.NewCompartmentButton.Location = new System.Drawing.Point(3, 505);
            this.NewCompartmentButton.Name = "NewCompartmentButton";
            this.NewCompartmentButton.Size = new System.Drawing.Size(144, 29);
            this.NewCompartmentButton.TabIndex = 2;
            this.NewCompartmentButton.Text = "New Compartment";
            this.NewCompartmentButton.UseVisualStyleBackColor = false;
            this.NewCompartmentButton.Click += NewCompartmentButton_Click;
            // 
            // BackArrowButton
            // 
            this.BackArrowButton.BackColor = System.Drawing.Color.Transparent;
            this.BackArrowButton.Cursor = System.Windows.Forms.Cursors.Default;
            this.BackArrowButton.FlatAppearance.BorderSize = 0;
            this.BackArrowButton.FlatAppearance.CheckedBackColor = System.Drawing.Color.Transparent;
            this.BackArrowButton.Image = global::SolidWorksTankDesign.Properties.Resources.left;
            this.BackArrowButton.Location = new System.Drawing.Point(59, 588);
            this.BackArrowButton.Name = "BackArrowButton";
            this.BackArrowButton.Size = new System.Drawing.Size(56, 36);
            this.BackArrowButton.TabIndex = 7;
            this.BackArrowButton.UseVisualStyleBackColor = false;
            this.BackArrowButton.Click += BackArrowButton_Click;
            // 
            // ForwardArrowButton
            // 
            this.ForwardArrowButton.BackColor = System.Drawing.Color.Transparent;
            this.ForwardArrowButton.Cursor = System.Windows.Forms.Cursors.Default;
            this.ForwardArrowButton.FlatAppearance.BorderSize = 0;
            this.ForwardArrowButton.FlatAppearance.CheckedBackColor = System.Drawing.Color.Transparent;
            this.ForwardArrowButton.Image = global::SolidWorksTankDesign.Properties.Resources.right;
            this.ForwardArrowButton.Location = new System.Drawing.Point(342, 588);
            this.ForwardArrowButton.Name = "ForwardArrowButton";
            this.ForwardArrowButton.Size = new System.Drawing.Size(56, 36);
            this.ForwardArrowButton.TabIndex = 6;
            this.ForwardArrowButton.UseVisualStyleBackColor = false;
            this.ForwardArrowButton.Click += new System.EventHandler(this.ForwardArrowButton_Click);
            // 
            // CompartmentsConfigurationLabel
            // 
            this.CompartmentsConfigurationLabel.AutoSize = true;
            this.CompartmentsConfigurationLabel.Location = new System.Drawing.Point(41, 9);
            this.CompartmentsConfigurationLabel.Name = "CompartmentsConfigurationLabel";
            this.CompartmentsConfigurationLabel.Size = new System.Drawing.Size(35, 13);
            this.CompartmentsConfigurationLabel.TabIndex = 8;
            this.CompartmentsConfigurationLabel.Text = "label1";
            this.CompartmentsConfigurationLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            // 
            // CompartmentWindowView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.CompartmentsConfigurationLabel);
            this.Controls.Add(this.BackArrowButton);
            this.Controls.Add(this.ForwardArrowButton);
            this.Controls.Add(this.NewCompartmentButton);
            this.Name = "CompartmentWindowView";
            this.Size = new System.Drawing.Size(459, 1032);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button NewCompartmentButton;
        private System.Windows.Forms.Button BackArrowButton;
        private System.Windows.Forms.Button ForwardArrowButton;
        private System.Windows.Forms.Label CompartmentsConfigurationLabel;
    }
}
