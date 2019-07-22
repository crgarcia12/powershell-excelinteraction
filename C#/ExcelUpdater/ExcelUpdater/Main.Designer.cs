namespace ExcelUpdater
{
    partial class Main
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
            this.btnOpenAndTransform = new System.Windows.Forms.Button();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.txtCopyUntil = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtInitialRowNr = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnOpenAndTransform
            // 
            this.btnOpenAndTransform.Location = new System.Drawing.Point(731, 87);
            this.btnOpenAndTransform.Name = "btnOpenAndTransform";
            this.btnOpenAndTransform.Size = new System.Drawing.Size(245, 40);
            this.btnOpenAndTransform.TabIndex = 0;
            this.btnOpenAndTransform.Text = "Open && Transform";
            this.btnOpenAndTransform.UseVisualStyleBackColor = true;
            this.btnOpenAndTransform.Click += new System.EventHandler(this.btnOpenAndTransform_Click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(174, 12);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(802, 26);
            this.txtFilePath.TabIndex = 1;
            // 
            // txtCopyUntil
            // 
            this.txtCopyUntil.Location = new System.Drawing.Point(174, 76);
            this.txtCopyUntil.Name = "txtCopyUntil";
            this.txtCopyUntil.Size = new System.Drawing.Size(115, 26);
            this.txtCopyUntil.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 20);
            this.label1.TabIndex = 3;
            this.label1.Text = "File full path";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(124, 20);
            this.label2.TabIndex = 4;
            this.label2.Text = "Copy until Max +";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(23, 47);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(145, 20);
            this.label3.TabIndex = 6;
            this.label3.Text = "Initial Table Row Nr";
            // 
            // txtInitialRowNr
            // 
            this.txtInitialRowNr.Location = new System.Drawing.Point(174, 44);
            this.txtInitialRowNr.Name = "txtInitialRowNr";
            this.txtInitialRowNr.Size = new System.Drawing.Size(115, 26);
            this.txtInitialRowNr.TabIndex = 5;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(988, 144);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtInitialRowNr);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtCopyUntil);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.btnOpenAndTransform);
            this.Name = "Main";
            this.Text = "Main";
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpenAndTransform;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.TextBox txtCopyUntil;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtInitialRowNr;
    }
}