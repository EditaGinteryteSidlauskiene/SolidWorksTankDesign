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
        private void InitializeComponent()
        {
            this.submitButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.numberOfDishedEnds = new System.Windows.Forms.NumericUpDown();
            this.numberOfCylindricalShells = new System.Windows.Forms.NumericUpDown();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numberOfDishedEnds)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numberOfCylindricalShells)).BeginInit();
            this.SuspendLayout();
            // 
            // submitButton
            // 
            this.submitButton.Location = new System.Drawing.Point(161, 217);
            this.submitButton.Name = "submitButton";
            this.submitButton.Size = new System.Drawing.Size(92, 39);
            this.submitButton.TabIndex = 0;
            this.submitButton.Text = "Submit";
            this.submitButton.UseVisualStyleBackColor = true;
            this.submitButton.Click += new System.EventHandler(this.SubmitButton);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(41, 116);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(133, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Set number of dished ends";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(41, 166);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(151, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Set number of cylindrical shells";
            // 
            // numberOfDishedEnds
            // 
            this.numberOfDishedEnds.AccessibleName = "numberOfDishedEnds";
            this.numberOfDishedEnds.Location = new System.Drawing.Point(197, 114);
            this.numberOfDishedEnds.Name = "numberOfDishedEnds";
            this.numberOfDishedEnds.Size = new System.Drawing.Size(44, 20);
            this.numberOfDishedEnds.TabIndex = 3;
            // 
            // numberOfCylindricalShells
            // 
            this.numberOfCylindricalShells.AccessibleName = "NumberOfCylindricalShells";
            this.numberOfCylindricalShells.Location = new System.Drawing.Point(226, 166);
            this.numberOfCylindricalShells.Name = "numberOfCylindricalShells";
            this.numberOfCylindricalShells.Size = new System.Drawing.Size(44, 20);
            this.numberOfCylindricalShells.TabIndex = 4;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(91, 311);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(150, 43);
            this.button1.TabIndex = 5;
            this.button1.Text = "Test";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // TaskpaneHostUI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.numberOfCylindricalShells);
            this.Controls.Add(this.numberOfDishedEnds);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.submitButton);
            this.Name = "TaskpaneHostUI";
            this.Size = new System.Drawing.Size(538, 972);
            ((System.ComponentModel.ISupportInitialize)(this.numberOfDishedEnds)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numberOfCylindricalShells)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button submitButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown numberOfDishedEnds;
        private System.Windows.Forms.NumericUpDown numberOfCylindricalShells;
        private System.Windows.Forms.Button button1;
    }
}
