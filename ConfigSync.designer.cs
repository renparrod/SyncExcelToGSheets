
namespace SyncExcelToGSheets
{
    partial class ConfigSync
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
            this.OkButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.GoogleSheetIdLabel = new System.Windows.Forms.Label();
            this.TextGoogleSheetId = new System.Windows.Forms.TextBox();
            this.LocalSheet = new System.Windows.Forms.Label();
            this.DropExcelSheetName = new System.Windows.Forms.ComboBox();
            this.TextGoogleSheetName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // OkButton
            // 
            this.OkButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.OkButton.Location = new System.Drawing.Point(410, 202);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(75, 23);
            this.OkButton.TabIndex = 0;
            this.OkButton.Text = "&Aceptar";
            this.OkButton.UseVisualStyleBackColor = true;
            this.OkButton.Click += new System.EventHandler(this.OkButton_Click);
            // 
            // CancelButton
            // 
            this.CancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelButton.Location = new System.Drawing.Point(491, 202);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(75, 23);
            this.CancelButton.TabIndex = 1;
            this.CancelButton.Text = "&Cancelar";
            this.CancelButton.UseVisualStyleBackColor = true;
            // 
            // GoogleSheetIdLabel
            // 
            this.GoogleSheetIdLabel.AutoSize = true;
            this.GoogleSheetIdLabel.Location = new System.Drawing.Point(12, 26);
            this.GoogleSheetIdLabel.Name = "GoogleSheetIdLabel";
            this.GoogleSheetIdLabel.Size = new System.Drawing.Size(83, 13);
            this.GoogleSheetIdLabel.TabIndex = 2;
            this.GoogleSheetIdLabel.Text = "Google Sheet id";
            // 
            // TextGoogleSheetId
            // 
            this.TextGoogleSheetId.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TextGoogleSheetId.Location = new System.Drawing.Point(165, 23);
            this.TextGoogleSheetId.Name = "TextGoogleSheetId";
            this.TextGoogleSheetId.Size = new System.Drawing.Size(401, 20);
            this.TextGoogleSheetId.TabIndex = 3;
            // 
            // LocalSheet
            // 
            this.LocalSheet.AutoSize = true;
            this.LocalSheet.Location = new System.Drawing.Point(12, 52);
            this.LocalSheet.Name = "LocalSheet";
            this.LocalSheet.Size = new System.Drawing.Size(96, 13);
            this.LocalSheet.TabIndex = 4;
            this.LocalSheet.Text = "Nombre hoja Excel";
            // 
            // DropExcelSheetName
            // 
            this.DropExcelSheetName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DropExcelSheetName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.DropExcelSheetName.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.DropExcelSheetName.FormattingEnabled = true;
            this.DropExcelSheetName.Location = new System.Drawing.Point(165, 49);
            this.DropExcelSheetName.Name = "DropExcelSheetName";
            this.DropExcelSheetName.Size = new System.Drawing.Size(401, 21);
            this.DropExcelSheetName.TabIndex = 5;
            // 
            // TextGoogleSheetName
            // 
            this.TextGoogleSheetName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TextGoogleSheetName.Location = new System.Drawing.Point(165, 76);
            this.TextGoogleSheetName.Name = "TextGoogleSheetName";
            this.TextGoogleSheetName.Size = new System.Drawing.Size(401, 20);
            this.TextGoogleSheetName.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 79);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(135, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Nombre hoja Google Sheet";
            // 
            // ConfigSync
            // 
            this.AcceptButton = this.OkButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(578, 237);
            this.Controls.Add(this.TextGoogleSheetName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.DropExcelSheetName);
            this.Controls.Add(this.LocalSheet);
            this.Controls.Add(this.TextGoogleSheetId);
            this.Controls.Add(this.GoogleSheetIdLabel);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.OkButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ConfigSync";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Configurar connexion a Google Sheets";
            this.Load += new System.EventHandler(this.ConfigGoogleSheet_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Button OkButton;
        private System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.Label GoogleSheetIdLabel;
        private System.Windows.Forms.TextBox TextGoogleSheetId;
        private System.Windows.Forms.Label LocalSheet;
        private System.Windows.Forms.ComboBox DropExcelSheetName;
        private System.Windows.Forms.TextBox TextGoogleSheetName;
        private System.Windows.Forms.Label label1;

        #endregion
    }
}