namespace desktopDashboard___Y_Lee.Forms
{
    partial class deactivatedUser
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
            this.pnDeactivatedUserMain = new System.Windows.Forms.Panel();
            this.comboxDeactivatedUser = new System.Windows.Forms.ComboBox();
            this.lbDeactivatedUserDropBox = new System.Windows.Forms.Label();
            this.btnDeactivatedUser = new System.Windows.Forms.Button();
            this.pnDeactivatedUserTop = new System.Windows.Forms.Panel();
            this.lbDeactivatedUserTop = new System.Windows.Forms.Label();
            this.rtxtDeactivatedUser = new System.Windows.Forms.RichTextBox();
            this.pnDeactivatedUserMain.SuspendLayout();
            this.pnDeactivatedUserTop.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnDeactivatedUserMain
            // 
            this.pnDeactivatedUserMain.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(42)))), ((int)(((byte)(64)))));
            this.pnDeactivatedUserMain.Controls.Add(this.comboxDeactivatedUser);
            this.pnDeactivatedUserMain.Controls.Add(this.lbDeactivatedUserDropBox);
            this.pnDeactivatedUserMain.Controls.Add(this.btnDeactivatedUser);
            this.pnDeactivatedUserMain.Controls.Add(this.pnDeactivatedUserTop);
            this.pnDeactivatedUserMain.Controls.Add(this.rtxtDeactivatedUser);
            this.pnDeactivatedUserMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnDeactivatedUserMain.Location = new System.Drawing.Point(0, 0);
            this.pnDeactivatedUserMain.Margin = new System.Windows.Forms.Padding(2);
            this.pnDeactivatedUserMain.Name = "pnDeactivatedUserMain";
            this.pnDeactivatedUserMain.Size = new System.Drawing.Size(901, 680);
            this.pnDeactivatedUserMain.TabIndex = 3;
            // 
            // comboxDeactivatedUser
            // 
            this.comboxDeactivatedUser.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboxDeactivatedUser.FormattingEnabled = true;
            this.comboxDeactivatedUser.Items.AddRange(new object[] {
            "MVW",
            "BMC",
            "CHO",
            "LAR",
            "HDC"});
            this.comboxDeactivatedUser.Location = new System.Drawing.Point(26, 62);
            this.comboxDeactivatedUser.Name = "comboxDeactivatedUser";
            this.comboxDeactivatedUser.Size = new System.Drawing.Size(121, 21);
            this.comboxDeactivatedUser.TabIndex = 16;
            // 
            // lbDeactivatedUserDropBox
            // 
            this.lbDeactivatedUserDropBox.AutoSize = true;
            this.lbDeactivatedUserDropBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbDeactivatedUserDropBox.ForeColor = System.Drawing.SystemColors.Window;
            this.lbDeactivatedUserDropBox.Location = new System.Drawing.Point(23, 42);
            this.lbDeactivatedUserDropBox.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbDeactivatedUserDropBox.Name = "lbDeactivatedUserDropBox";
            this.lbDeactivatedUserDropBox.Size = new System.Drawing.Size(71, 16);
            this.lbDeactivatedUserDropBox.TabIndex = 15;
            this.lbDeactivatedUserDropBox.Text = "Select Site";
            // 
            // btnDeactivatedUser
            // 
            this.btnDeactivatedUser.Location = new System.Drawing.Point(165, 62);
            this.btnDeactivatedUser.Margin = new System.Windows.Forms.Padding(2);
            this.btnDeactivatedUser.Name = "btnDeactivatedUser";
            this.btnDeactivatedUser.Size = new System.Drawing.Size(56, 21);
            this.btnDeactivatedUser.TabIndex = 7;
            this.btnDeactivatedUser.Text = "Display";
            this.btnDeactivatedUser.UseVisualStyleBackColor = true;
            this.btnDeactivatedUser.Click += new System.EventHandler(this.btnDeactivatedUser_Click);
            // 
            // pnDeactivatedUserTop
            // 
            this.pnDeactivatedUserTop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(42)))), ((int)(((byte)(100)))));
            this.pnDeactivatedUserTop.Controls.Add(this.lbDeactivatedUserTop);
            this.pnDeactivatedUserTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnDeactivatedUserTop.Location = new System.Drawing.Point(0, 0);
            this.pnDeactivatedUserTop.Margin = new System.Windows.Forms.Padding(2);
            this.pnDeactivatedUserTop.Name = "pnDeactivatedUserTop";
            this.pnDeactivatedUserTop.Size = new System.Drawing.Size(901, 34);
            this.pnDeactivatedUserTop.TabIndex = 2;
            // 
            // lbDeactivatedUserTop
            // 
            this.lbDeactivatedUserTop.AutoSize = true;
            this.lbDeactivatedUserTop.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbDeactivatedUserTop.ForeColor = System.Drawing.Color.White;
            this.lbDeactivatedUserTop.Location = new System.Drawing.Point(357, 9);
            this.lbDeactivatedUserTop.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lbDeactivatedUserTop.Name = "lbDeactivatedUserTop";
            this.lbDeactivatedUserTop.Size = new System.Drawing.Size(228, 20);
            this.lbDeactivatedUserTop.TabIndex = 7;
            this.lbDeactivatedUserTop.Text = "Account Deactivated Users";
            // 
            // rtxtDeactivatedUser
            // 
            this.rtxtDeactivatedUser.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(33)))), ((int)(((byte)(36)))));
            this.rtxtDeactivatedUser.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rtxtDeactivatedUser.ForeColor = System.Drawing.Color.White;
            this.rtxtDeactivatedUser.Location = new System.Drawing.Point(22, 88);
            this.rtxtDeactivatedUser.Margin = new System.Windows.Forms.Padding(2);
            this.rtxtDeactivatedUser.Name = "rtxtDeactivatedUser";
            this.rtxtDeactivatedUser.Size = new System.Drawing.Size(856, 524);
            this.rtxtDeactivatedUser.TabIndex = 1;
            this.rtxtDeactivatedUser.Text = "";
            // 
            // deactivatedUser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(901, 680);
            this.Controls.Add(this.pnDeactivatedUserMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "deactivatedUser";
            this.Text = "createUser";
            this.pnDeactivatedUserMain.ResumeLayout(false);
            this.pnDeactivatedUserMain.PerformLayout();
            this.pnDeactivatedUserTop.ResumeLayout(false);
            this.pnDeactivatedUserTop.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnDeactivatedUserMain;
        private System.Windows.Forms.Panel pnDeactivatedUserTop;
        private System.Windows.Forms.Label lbDeactivatedUserTop;
        private System.Windows.Forms.RichTextBox rtxtDeactivatedUser;
        private System.Windows.Forms.Button btnDeactivatedUser;
        private System.Windows.Forms.ComboBox comboxDeactivatedUser;
        private System.Windows.Forms.Label lbDeactivatedUserDropBox;
    }
}