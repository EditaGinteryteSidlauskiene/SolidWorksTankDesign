namespace SolidWorksTankDesign.MVP.Views
{
    partial class PaintingSystemsManagerView
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
            this.TreatmentDescriptionPanel = new System.Windows.Forms.Panel();
            this.EditTreatmentCheckBox = new System.Windows.Forms.CheckBox();
            this.SaveButton = new System.Windows.Forms.Button();
            this.DeleteTreatmentButton = new System.Windows.Forms.Button();
            this.TreatmentCommentsTextBox = new System.Windows.Forms.TextBox();
            this.CommentsLabel = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.TreatmentCleaningTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.TreatmentDescriptionTextBox = new System.Windows.Forms.TextBox();
            this.TreatmentDescriptionLabel = new System.Windows.Forms.Label();
            this.TreatmentsListBox = new System.Windows.Forms.ListBox();
            this.TreatmentDescriptionPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // BackArrowButton
            // 
            this.BackArrowButton.BackColor = System.Drawing.Color.Transparent;
            this.BackArrowButton.Cursor = System.Windows.Forms.Cursors.Default;
            this.BackArrowButton.FlatAppearance.BorderSize = 0;
            this.BackArrowButton.FlatAppearance.CheckedBackColor = System.Drawing.Color.Transparent;
            this.BackArrowButton.Image = global::SolidWorksTankDesign.Properties.Resources.left;
            this.BackArrowButton.Location = new System.Drawing.Point(59, 789);
            this.BackArrowButton.Name = "BackArrowButton";
            this.BackArrowButton.Size = new System.Drawing.Size(56, 36);
            this.BackArrowButton.TabIndex = 8;
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
            this.ForwardArrowButton.Location = new System.Drawing.Point(342, 789);
            this.ForwardArrowButton.Name = "ForwardArrowButton";
            this.ForwardArrowButton.Size = new System.Drawing.Size(56, 36);
            this.ForwardArrowButton.TabIndex = 9;
            this.ForwardArrowButton.UseVisualStyleBackColor = false;
            // 
            // TreatmentDescriptionPanel
            // 
            this.TreatmentDescriptionPanel.Controls.Add(this.EditTreatmentCheckBox);
            this.TreatmentDescriptionPanel.Controls.Add(this.SaveButton);
            this.TreatmentDescriptionPanel.Controls.Add(this.DeleteTreatmentButton);
            this.TreatmentDescriptionPanel.Controls.Add(this.TreatmentCommentsTextBox);
            this.TreatmentDescriptionPanel.Controls.Add(this.CommentsLabel);
            this.TreatmentDescriptionPanel.Controls.Add(this.label2);
            this.TreatmentDescriptionPanel.Controls.Add(this.TreatmentCleaningTextBox);
            this.TreatmentDescriptionPanel.Controls.Add(this.label1);
            this.TreatmentDescriptionPanel.Controls.Add(this.TreatmentDescriptionTextBox);
            this.TreatmentDescriptionPanel.Controls.Add(this.TreatmentDescriptionLabel);
            this.TreatmentDescriptionPanel.Location = new System.Drawing.Point(1, 239);
            this.TreatmentDescriptionPanel.Name = "TreatmentDescriptionPanel";
            this.TreatmentDescriptionPanel.Size = new System.Drawing.Size(443, 523);
            this.TreatmentDescriptionPanel.TabIndex = 10;
            this.TreatmentDescriptionPanel.Visible = false;
            // 
            // EditTreatmentCheckBox
            // 
            this.EditTreatmentCheckBox.AutoSize = true;
            this.EditTreatmentCheckBox.CheckAlign = System.Drawing.ContentAlignment.TopRight;
            this.EditTreatmentCheckBox.Location = new System.Drawing.Point(382, 16);
            this.EditTreatmentCheckBox.Name = "EditTreatmentCheckBox";
            this.EditTreatmentCheckBox.Size = new System.Drawing.Size(44, 17);
            this.EditTreatmentCheckBox.TabIndex = 14;
            this.EditTreatmentCheckBox.Text = "Edit";
            this.EditTreatmentCheckBox.UseVisualStyleBackColor = true;
            this.EditTreatmentCheckBox.CheckedChanged += new System.EventHandler(this.EditTreatmentCheckBox_CheckedChanged);
            // 
            // SaveButton
            // 
            this.SaveButton.Location = new System.Drawing.Point(100, 209);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(99, 23);
            this.SaveButton.TabIndex = 12;
            this.SaveButton.Text = "Save Treatment";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += SaveButton_Click;
            // 
            // DeleteTreatmentButton
            // 
            this.DeleteTreatmentButton.Location = new System.Drawing.Point(250, 209);
            this.DeleteTreatmentButton.Name = "DeleteTreatmentButton";
            this.DeleteTreatmentButton.Size = new System.Drawing.Size(99, 25);
            this.DeleteTreatmentButton.TabIndex = 10;
            this.DeleteTreatmentButton.Text = "Delete Treatment";
            this.DeleteTreatmentButton.UseVisualStyleBackColor = true;
            this.DeleteTreatmentButton.Click += DeleteTreatmentButton_Click;
            // 
            // TreatmentCommentsTextBox
            // 
            this.TreatmentCommentsTextBox.AcceptsReturn = true;
            this.TreatmentCommentsTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TreatmentCommentsTextBox.Location = new System.Drawing.Point(9, 271);
            this.TreatmentCommentsTextBox.Multiline = true;
            this.TreatmentCommentsTextBox.Name = "TreatmentCommentsTextBox";
            this.TreatmentCommentsTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.TreatmentCommentsTextBox.Size = new System.Drawing.Size(433, 20);
            this.TreatmentCommentsTextBox.TabIndex = 7;
            this.TreatmentCommentsTextBox.TextChanged += TreatmentCommentsTextBox_TextChanged;
            // 
            // CommentsLabel
            // 
            this.CommentsLabel.AutoSize = true;
            this.CommentsLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CommentsLabel.Location = new System.Drawing.Point(11, 253);
            this.CommentsLabel.Name = "CommentsLabel";
            this.CommentsLabel.Size = new System.Drawing.Size(75, 15);
            this.CommentsLabel.TabIndex = 6;
            this.CommentsLabel.Text = "Comments";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(4, 83);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(102, 15);
            this.label2.TabIndex = 4;
            this.label2.Text = "Coating Layers";
            // 
            // TreatmentCleaningTextBox
            // 
            this.TreatmentCleaningTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TreatmentCleaningTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TreatmentCleaningTextBox.Location = new System.Drawing.Point(113, 44);
            this.TreatmentCleaningTextBox.Name = "TreatmentCleaningTextBox";
            this.TreatmentCleaningTextBox.Size = new System.Drawing.Size(200, 21);
            this.TreatmentCleaningTextBox.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(4, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "Cleaning";
            // 
            // TreatmentDescriptionTextBox
            // 
            this.TreatmentDescriptionTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TreatmentDescriptionTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TreatmentDescriptionTextBox.Location = new System.Drawing.Point(113, 12);
            this.TreatmentDescriptionTextBox.Name = "TreatmentDescriptionTextBox";
            this.TreatmentDescriptionTextBox.Size = new System.Drawing.Size(200, 21);
            this.TreatmentDescriptionTextBox.TabIndex = 1;
            // 
            // TreatmentDescriptionLabel
            // 
            this.TreatmentDescriptionLabel.AutoSize = true;
            this.TreatmentDescriptionLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TreatmentDescriptionLabel.Location = new System.Drawing.Point(4, 13);
            this.TreatmentDescriptionLabel.Name = "TreatmentDescriptionLabel";
            this.TreatmentDescriptionLabel.Size = new System.Drawing.Size(80, 15);
            this.TreatmentDescriptionLabel.TabIndex = 0;
            this.TreatmentDescriptionLabel.Text = "Description";
            // 
            // TreatmentsListBox
            // 
            this.TreatmentsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TreatmentsListBox.FormattingEnabled = true;
            this.TreatmentsListBox.ItemHeight = 15;
            this.TreatmentsListBox.Location = new System.Drawing.Point(4, 4);
            this.TreatmentsListBox.Name = "TreatmentsListBox";
            this.TreatmentsListBox.Size = new System.Drawing.Size(440, 229);
            this.TreatmentsListBox.TabIndex = 11;
            this.TreatmentsListBox.SelectedValueChanged += TreatmentsListBox_SelectedValueChanged;
            // 
            // PaintingSystemsManagerView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.TreatmentsListBox);
            this.Controls.Add(this.TreatmentDescriptionPanel);
            this.Controls.Add(this.ForwardArrowButton);
            this.Controls.Add(this.BackArrowButton);
            this.Name = "PaintingSystemsManagerView";
            this.Size = new System.Drawing.Size(450, 1086);
            this.TreatmentDescriptionPanel.ResumeLayout(false);
            this.TreatmentDescriptionPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button BackArrowButton;
        private System.Windows.Forms.Button ForwardArrowButton;
        private System.Windows.Forms.Panel TreatmentDescriptionPanel;
        private System.Windows.Forms.Label TreatmentDescriptionLabel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TreatmentCleaningTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TreatmentDescriptionTextBox;
        private System.Windows.Forms.TextBox TreatmentCommentsTextBox;
        private System.Windows.Forms.Label CommentsLabel;
        private System.Windows.Forms.ListBox TreatmentsListBox;
        private System.Windows.Forms.Button DeleteTreatmentButton;
        private System.Windows.Forms.Button SaveButton;
        private System.Windows.Forms.CheckBox EditTreatmentCheckBox;
    }
}
