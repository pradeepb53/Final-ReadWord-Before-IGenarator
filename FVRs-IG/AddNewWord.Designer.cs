namespace FVRs_IG
{
    partial class AddNewWord
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
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxNewWord = new System.Windows.Forms.TextBox();
            this.buttonAddNewWord = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(64, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Enter New Word :";
            // 
            // textBoxNewWord
            // 
            this.textBoxNewWord.Location = new System.Drawing.Point(185, 54);
            this.textBoxNewWord.Name = "textBoxNewWord";
            this.textBoxNewWord.Size = new System.Drawing.Size(140, 20);
            this.textBoxNewWord.TabIndex = 1;
            this.textBoxNewWord.TextChanged += new System.EventHandler(this.textBoxNewWord_TextChanged);
            // 
            // buttonAddNewWord
            // 
            this.buttonAddNewWord.Location = new System.Drawing.Point(211, 116);
            this.buttonAddNewWord.Name = "buttonAddNewWord";
            this.buttonAddNewWord.Size = new System.Drawing.Size(75, 23);
            this.buttonAddNewWord.TabIndex = 2;
            this.buttonAddNewWord.Text = "Add";
            this.buttonAddNewWord.UseVisualStyleBackColor = true;
            this.buttonAddNewWord.Click += new System.EventHandler(this.buttonAddNewWord_Click);
            // 
            // AddNewWord
            // 
            this.AcceptButton = this.buttonAddNewWord;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(427, 171);
            this.Controls.Add(this.buttonAddNewWord);
            this.Controls.Add(this.textBoxNewWord);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AddNewWord";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Add - New Word";
            this.Load += new System.EventHandler(this.AddNewWord_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxNewWord;
        private System.Windows.Forms.Button buttonAddNewWord;
    }
}