using System;

namespace TIA_Add_In_Example_Project
{


    partial class Form
    {

        public int adressStart => (int)numericUpDownStarting.Value;
        public int addresSize => (int)numericUpDownSize.Value;

        public int addressIncr => (int)numericUpDownIncrement.Value;



        private void numericUpDownStarting_Enter(object sender, EventArgs e)
        {
            numericUpDownStarting.Select(0, numericUpDownStarting.Text.Length);
        }

        private void numericUpDownIncrement_Enter(object sender, EventArgs e)
        {
            numericUpDownIncrement.Select(0, numericUpDownIncrement.Text.Length);
        }

        private void numericUpDownSize_Enter(object sender, EventArgs e)
        {
            numericUpDownSize.Select(0, numericUpDownSize.Text.Length);
        }

        private void numericUpDownType_Enter(object sender, EventArgs e)
        {
            numericUpDownSize.Select(0, numericUpDownSize.Text.Length);
        }

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form));
            this.numericUpDownStarting = new System.Windows.Forms.NumericUpDown();
            this.numericUpDownIncrement = new System.Windows.Forms.NumericUpDown();
            this.numericUpDownSize = new System.Windows.Forms.NumericUpDown();
            this.labelStartingNumber = new System.Windows.Forms.Label();
            this.labelIncrement = new System.Windows.Forms.Label();
            this.labelSize = new System.Windows.Forms.Label();
            this.buttonOk = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStarting)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownIncrement)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownSize)).BeginInit();
            this.SuspendLayout();
            // 
            // numericUpDownStarting
            // 
            this.numericUpDownStarting.BackColor = System.Drawing.Color.White;
            this.numericUpDownStarting.Location = new System.Drawing.Point(150, 32);
            this.numericUpDownStarting.Maximum = new decimal(new int[] {
            59999,
            0,
            0,
            0});
            this.numericUpDownStarting.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownStarting.Name = "numericUpDownStarting";
            this.numericUpDownStarting.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.numericUpDownStarting.Size = new System.Drawing.Size(120, 20);
            this.numericUpDownStarting.TabIndex = 0;
            this.numericUpDownStarting.Value = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numericUpDownStarting.Enter += new System.EventHandler(this.numericUpDownStarting_Enter);
            // 
            // numericUpDownIncrement
            // 
            this.numericUpDownIncrement.BackColor = System.Drawing.Color.White;
            this.numericUpDownIncrement.Location = new System.Drawing.Point(150, 62);
            this.numericUpDownIncrement.Maximum = new decimal(new int[] {
            59999,
            0,
            0,
            0});
            this.numericUpDownIncrement.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownIncrement.Name = "numericUpDownIncrement";
            this.numericUpDownIncrement.Size = new System.Drawing.Size(120, 20);
            this.numericUpDownIncrement.TabIndex = 1;
            this.numericUpDownIncrement.Value = new decimal(new int[] {
            24,
            0,
            0,
            0});
            this.numericUpDownIncrement.Enter += new System.EventHandler(this.numericUpDownIncrement_Enter);
            // 
            // numericUpDownSize
            // 
            this.numericUpDownSize.BackColor = System.Drawing.Color.White;
            this.numericUpDownSize.Location = new System.Drawing.Point(150, 92);
            this.numericUpDownSize.Maximum = new decimal(new int[] {
            59999,
            0,
            0,
            0});
            this.numericUpDownSize.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownSize.Name = "numericUpDownSize";
            this.numericUpDownSize.Size = new System.Drawing.Size(120, 20);
            this.numericUpDownSize.TabIndex = 0;
            this.numericUpDownSize.Value = new decimal(new int[] {
            12,
            0,
            0,
            0});
            // 
            // labelStartingNumber
            // 
            this.labelStartingNumber.AutoSize = true;
            this.labelStartingNumber.Location = new System.Drawing.Point(10, 32);
            this.labelStartingNumber.Name = "labelStartingNumber";
            this.labelStartingNumber.Size = new System.Drawing.Size(87, 13);
            this.labelStartingNumber.TabIndex = 2;
            this.labelStartingNumber.Text = "Starting Address:";
            // 
            // labelIncrement
            // 
            this.labelIncrement.AutoSize = true;
            this.labelIncrement.Location = new System.Drawing.Point(10, 62);
            this.labelIncrement.Name = "labelIncrement";
            this.labelIncrement.Size = new System.Drawing.Size(97, 13);
            this.labelIncrement.TabIndex = 3;
            this.labelIncrement.Text = "Address increment:";
            // 
            // labelSize
            // 
            this.labelSize.AutoSize = true;
            this.labelSize.Location = new System.Drawing.Point(10, 92);
            this.labelSize.Name = "labelSize";
            this.labelSize.Size = new System.Drawing.Size(75, 13);
            this.labelSize.TabIndex = 3;
            this.labelSize.Text = "Telegram size:";
            // 
            // buttonOk
            // 
            this.buttonOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonOk.Location = new System.Drawing.Point(148, 157);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(75, 23);
            this.buttonOk.TabIndex = 4;
            this.buttonOk.Text = "Ok";
            this.buttonOk.UseVisualStyleBackColor = true;
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(245, 157);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 5;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(473, 211);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.labelIncrement);
            this.Controls.Add(this.labelStartingNumber);
            this.Controls.Add(this.labelSize);
            this.Controls.Add(this.numericUpDownIncrement);
            this.Controls.Add(this.numericUpDownSize);
            this.Controls.Add(this.numericUpDownStarting);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Set telegram";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStarting)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownIncrement)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownSize)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }



        public System.Windows.Forms.NumericUpDown numericUpDownStarting;
        public System.Windows.Forms.NumericUpDown numericUpDownSize;
        public System.Windows.Forms.NumericUpDown numericUpDownIncrement;
        public System.Windows.Forms.Label labelStartingNumber;
        public System.Windows.Forms.Label labelSize;
        public System.Windows.Forms.Label labelIncrement;
        public System.Windows.Forms.Button buttonOk;
        public System.Windows.Forms.Button buttonCancel;

        #endregion
    }
}