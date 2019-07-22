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
            this.SuspendLayout();
            // 
            // btnOpenAndTransform
            // 
            this.btnOpenAndTransform.Location = new System.Drawing.Point(731, 44);
            this.btnOpenAndTransform.Name = "btnOpenAndTransform";
            this.btnOpenAndTransform.Size = new System.Drawing.Size(245, 40);
            this.btnOpenAndTransform.TabIndex = 0;
            this.btnOpenAndTransform.Text = "Open & Transform";
            this.btnOpenAndTransform.UseVisualStyleBackColor = true;
            this.btnOpenAndTransform.Click += new System.EventHandler(this.btnOpenAndTransform_Click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(12, 12);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(964, 26);
            this.txtFilePath.TabIndex = 1;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(988, 108);
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
    }
}