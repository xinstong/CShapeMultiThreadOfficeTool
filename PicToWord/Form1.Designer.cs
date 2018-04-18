namespace PicToWord
{
    partial class Form1
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
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.progressBarHandle = new System.Windows.Forms.ProgressBar();
            this.buttonStartHandle = new System.Windows.Forms.Button();
            this.listBoxResult = new System.Windows.Forms.ListBox();
            this.textBoxPath = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Location = new System.Drawing.Point(806, 8);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(75, 23);
            this.buttonBrowse.TabIndex = 0;
            this.buttonBrowse.Text = "浏览";
            this.buttonBrowse.UseVisualStyleBackColor = true;
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // progressBarHandle
            // 
            this.progressBarHandle.Location = new System.Drawing.Point(13, 36);
            this.progressBarHandle.Name = "progressBarHandle";
            this.progressBarHandle.Size = new System.Drawing.Size(787, 23);
            this.progressBarHandle.TabIndex = 2;
            // 
            // buttonStartHandle
            // 
            this.buttonStartHandle.Enabled = false;
            this.buttonStartHandle.Location = new System.Drawing.Point(806, 35);
            this.buttonStartHandle.Name = "buttonStartHandle";
            this.buttonStartHandle.Size = new System.Drawing.Size(75, 23);
            this.buttonStartHandle.TabIndex = 3;
            this.buttonStartHandle.Text = "开始处理";
            this.buttonStartHandle.UseVisualStyleBackColor = true;
            this.buttonStartHandle.Click += new System.EventHandler(this.buttonStartHandle_Click);
            // 
            // listBoxResult
            // 
            this.listBoxResult.FormattingEnabled = true;
            this.listBoxResult.Location = new System.Drawing.Point(13, 66);
            this.listBoxResult.Name = "listBoxResult";
            this.listBoxResult.Size = new System.Drawing.Size(868, 459);
            this.listBoxResult.TabIndex = 4;

            // 
            // textBoxPath
            // 
            this.textBoxPath.Location = new System.Drawing.Point(12, 11);
            this.textBoxPath.Name = "textBoxPath";
            this.textBoxPath.Size = new System.Drawing.Size(788, 20);
            this.textBoxPath.TabIndex = 5;
            this.textBoxPath.TextChanged += new System.EventHandler(this.textBoxPath_TextChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(882, 540);
            this.Controls.Add(this.textBoxPath);
            this.Controls.Add(this.listBoxResult);
            this.Controls.Add(this.buttonStartHandle);
            this.Controls.Add(this.progressBarHandle);
            this.Controls.Add(this.buttonBrowse);
            this.Name = "Form1";
            this.Text = "批量处理工具-插入图片到word文档最后";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.ProgressBar progressBarHandle;
        private System.Windows.Forms.Button buttonStartHandle;
        private System.Windows.Forms.ListBox listBoxResult;
        private System.Windows.Forms.TextBox textBoxPath;
    }
}

