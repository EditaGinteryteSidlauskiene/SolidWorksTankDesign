using System.Drawing;
using System.Windows.Forms;

namespace SolidWorksTankDesign.Windows
{
    partial class ManagePaintingSystemsWindow
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
            this.TreatmentsLabel = new System.Windows.Forms.Label();
            this.TreatmentsListBox = new System.Windows.Forms.ListBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.UpdateButton = new System.Windows.Forms.Button();
            this.NewTreatmentButton = new System.Windows.Forms.Button();
            this.DeleteTreatmentButtob = new System.Windows.Forms.Button();
            this.DeleteLayerButton = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.BackArrowButton = new System.Windows.Forms.Button();
            this.ForwardArrowButton = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // TreatmentsLabel
            // 
            this.TreatmentsLabel.AutoSize = true;
            this.TreatmentsLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TreatmentsLabel.Location = new System.Drawing.Point(16, 17);
            this.TreatmentsLabel.Name = "TreatmentsLabel";
            this.TreatmentsLabel.Size = new System.Drawing.Size(85, 16);
            this.TreatmentsLabel.TabIndex = 0;
            this.TreatmentsLabel.Text = "Treatments";
            // 
            // TreatmentsListBox
            // 
            this.TreatmentsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TreatmentsListBox.FormattingEnabled = true;
            this.TreatmentsListBox.ItemHeight = 15;
            this.TreatmentsListBox.Location = new System.Drawing.Point(19, 37);
            this.TreatmentsListBox.Name = "TreatmentsListBox";
            this.TreatmentsListBox.Size = new System.Drawing.Size(376, 169);
            this.TreatmentsListBox.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.UpdateButton);
            this.panel1.Controls.Add(this.NewTreatmentButton);
            this.panel1.Controls.Add(this.DeleteTreatmentButtob);
            this.panel1.Controls.Add(this.DeleteLayerButton);
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Location = new System.Drawing.Point(19, 213);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(376, 269);
            this.panel1.TabIndex = 2;
            // 
            // UpdateButton
            // 
            this.UpdateButton.Location = new System.Drawing.Point(299, 194);
            this.UpdateButton.Name = "UpdateButton";
            this.UpdateButton.Size = new System.Drawing.Size(75, 23);
            this.UpdateButton.TabIndex = 4;
            this.UpdateButton.Text = "&Update";
            this.UpdateButton.UseVisualStyleBackColor = true;
            // 
            // NewTreatmentButton
            // 
            this.NewTreatmentButton.Location = new System.Drawing.Point(209, 194);
            this.NewTreatmentButton.Name = "NewTreatmentButton";
            this.NewTreatmentButton.Size = new System.Drawing.Size(75, 23);
            this.NewTreatmentButton.TabIndex = 3;
            this.NewTreatmentButton.Text = "&New Treatment";
            this.NewTreatmentButton.UseVisualStyleBackColor = true;
            // 
            // DeleteTreatmentButtob
            // 
            this.DeleteTreatmentButtob.Location = new System.Drawing.Point(90, 194);
            this.DeleteTreatmentButtob.Name = "DeleteTreatmentButtob";
            this.DeleteTreatmentButtob.Size = new System.Drawing.Size(104, 23);
            this.DeleteTreatmentButtob.TabIndex = 2;
            this.DeleteTreatmentButtob.Text = "&Delete Treatment";
            this.DeleteTreatmentButtob.UseVisualStyleBackColor = true;
            // 
            // DeleteLayerButton
            // 
            this.DeleteLayerButton.Location = new System.Drawing.Point(0, 194);
            this.DeleteLayerButton.Name = "DeleteLayerButton";
            this.DeleteLayerButton.Size = new System.Drawing.Size(75, 23);
            this.DeleteLayerButton.TabIndex = 1;
            this.DeleteLayerButton.Text = "&Delete Layer";
            this.DeleteLayerButton.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(376, 173);
            this.dataGridView1.TabIndex = 0;
            // 
            // BackArrowButton
            // 
            this.BackArrowButton.BackColor = System.Drawing.Color.Transparent;
            this.BackArrowButton.Cursor = System.Windows.Forms.Cursors.Default;
            this.BackArrowButton.FlatAppearance.BorderSize = 0;
            this.BackArrowButton.FlatAppearance.CheckedBackColor = System.Drawing.Color.Transparent;
            this.BackArrowButton.Image = global::SolidWorksTankDesign.Properties.Resources.left;
            this.BackArrowButton.Location = new System.Drawing.Point(61, 533);
            this.BackArrowButton.Name = "BackArrowButton";
            this.BackArrowButton.Size = new System.Drawing.Size(56, 36);
            this.BackArrowButton.TabIndex = 5;
            this.BackArrowButton.UseVisualStyleBackColor = false;
            // 
            // ForwardArrowButton
            // 
            this.ForwardArrowButton.BackColor = System.Drawing.Color.Transparent;
            this.ForwardArrowButton.Cursor = System.Windows.Forms.Cursors.Default;
            this.ForwardArrowButton.FlatAppearance.BorderSize = 0;
            this.ForwardArrowButton.FlatAppearance.CheckedBackColor = System.Drawing.Color.Transparent;
            this.ForwardArrowButton.Image = global::SolidWorksTankDesign.Properties.Resources.right;
            this.ForwardArrowButton.Location = new System.Drawing.Point(281, 533);
            this.ForwardArrowButton.Name = "ForwardArrowButton";
            this.ForwardArrowButton.Size = new System.Drawing.Size(56, 36);
            this.ForwardArrowButton.TabIndex = 4;
            this.ForwardArrowButton.UseVisualStyleBackColor = false;
            // 
            // ManagePaintingSystemsWindow
            // 
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.BackArrowButton);
            this.Controls.Add(this.ForwardArrowButton);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.TreatmentsListBox);
            this.Controls.Add(this.TreatmentsLabel);
            this.Name = "ManagePaintingSystemsWindow";
            this.Size = new System.Drawing.Size(426, 1031);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private Label TreatmentsLabel;
        private ListBox TreatmentsListBox;
        private Panel panel1;
        private DataGridView dataGridView1;
        private Button UpdateButton;
        private Button NewTreatmentButton;
        private Button DeleteTreatmentButtob;
        private Button DeleteLayerButton;
        private Button ForwardArrowButton;
        private Button BackArrowButton;
    }
}
