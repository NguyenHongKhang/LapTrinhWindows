
namespace WindowsFormsApp1.Views
{
    partial class AddAccount
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
            this.panel29 = new System.Windows.Forms.Panel();
            this.panel46 = new System.Windows.Forms.Panel();
            this.cbMaNhanVien = new System.Windows.Forms.ComboBox();
            this.label17 = new System.Windows.Forms.Label();
            this.panel38 = new System.Windows.Forms.Panel();
            this.cbLoaiTaiKhoan = new System.Windows.Forms.ComboBox();
            this.label16 = new System.Windows.Forms.Label();
            this.panel34 = new System.Windows.Forms.Panel();
            this.tbMatKhau = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.panel35 = new System.Windows.Forms.Panel();
            this.tbTenHienThi = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.panel36 = new System.Windows.Forms.Panel();
            this.tbTenTaiKhoan = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.bCancel = new System.Windows.Forms.Button();
            this.bOK = new System.Windows.Forms.Button();
            this.panel29.SuspendLayout();
            this.panel46.SuspendLayout();
            this.panel38.SuspendLayout();
            this.panel34.SuspendLayout();
            this.panel35.SuspendLayout();
            this.panel36.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel29
            // 
            this.panel29.Controls.Add(this.panel46);
            this.panel29.Controls.Add(this.panel38);
            this.panel29.Controls.Add(this.panel34);
            this.panel29.Controls.Add(this.panel35);
            this.panel29.Controls.Add(this.panel36);
            this.panel29.Location = new System.Drawing.Point(12, 12);
            this.panel29.Name = "panel29";
            this.panel29.Size = new System.Drawing.Size(493, 324);
            this.panel29.TabIndex = 4;
            // 
            // panel46
            // 
            this.panel46.Controls.Add(this.cbMaNhanVien);
            this.panel46.Controls.Add(this.label17);
            this.panel46.Location = new System.Drawing.Point(3, 249);
            this.panel46.Name = "panel46";
            this.panel46.Size = new System.Drawing.Size(483, 56);
            this.panel46.TabIndex = 6;
            // 
            // cbMaNhanVien
            // 
            this.cbMaNhanVien.FormattingEnabled = true;
            this.cbMaNhanVien.Location = new System.Drawing.Point(209, 14);
            this.cbMaNhanVien.Name = "cbMaNhanVien";
            this.cbMaNhanVien.Size = new System.Drawing.Size(254, 24);
            this.cbMaNhanVien.TabIndex = 1;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label17.Location = new System.Drawing.Point(3, 14);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(182, 29);
            this.label17.TabIndex = 0;
            this.label17.Text = "Mã Nhân Viên:";
            // 
            // panel38
            // 
            this.panel38.Controls.Add(this.cbLoaiTaiKhoan);
            this.panel38.Controls.Add(this.label16);
            this.panel38.Location = new System.Drawing.Point(3, 187);
            this.panel38.Name = "panel38";
            this.panel38.Size = new System.Drawing.Size(483, 56);
            this.panel38.TabIndex = 4;
            // 
            // cbLoaiTaiKhoan
            // 
            this.cbLoaiTaiKhoan.FormattingEnabled = true;
            this.cbLoaiTaiKhoan.Location = new System.Drawing.Point(209, 14);
            this.cbLoaiTaiKhoan.Name = "cbLoaiTaiKhoan";
            this.cbLoaiTaiKhoan.Size = new System.Drawing.Size(254, 24);
            this.cbLoaiTaiKhoan.TabIndex = 1;
            this.cbLoaiTaiKhoan.SelectedIndexChanged += new System.EventHandler(this.cbLoaiTaiKhoan_SelectedIndexChanged);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label16.Location = new System.Drawing.Point(3, 14);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(196, 29);
            this.label16.TabIndex = 0;
            this.label16.Text = "Loại Tài Khoản:";
            // 
            // panel34
            // 
            this.panel34.Controls.Add(this.tbMatKhau);
            this.panel34.Controls.Add(this.label4);
            this.panel34.Location = new System.Drawing.Point(3, 63);
            this.panel34.Name = "panel34";
            this.panel34.Size = new System.Drawing.Size(483, 56);
            this.panel34.TabIndex = 5;
            // 
            // tbMatKhau
            // 
            this.tbMatKhau.Location = new System.Drawing.Point(209, 14);
            this.tbMatKhau.Name = "tbMatKhau";
            this.tbMatKhau.Size = new System.Drawing.Size(254, 22);
            this.tbMatKhau.TabIndex = 2;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label4.Location = new System.Drawing.Point(3, 14);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(128, 29);
            this.label4.TabIndex = 0;
            this.label4.Text = "Mật Khẩu:";
            // 
            // panel35
            // 
            this.panel35.Controls.Add(this.tbTenHienThi);
            this.panel35.Controls.Add(this.label7);
            this.panel35.Location = new System.Drawing.Point(3, 125);
            this.panel35.Name = "panel35";
            this.panel35.Size = new System.Drawing.Size(483, 56);
            this.panel35.TabIndex = 2;
            // 
            // tbTenHienThi
            // 
            this.tbTenHienThi.Location = new System.Drawing.Point(209, 14);
            this.tbTenHienThi.Name = "tbTenHienThi";
            this.tbTenHienThi.Size = new System.Drawing.Size(254, 22);
            this.tbTenHienThi.TabIndex = 1;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label7.Location = new System.Drawing.Point(3, 14);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(172, 29);
            this.label7.TabIndex = 0;
            this.label7.Text = "Tên Hiển Thị:";
            // 
            // panel36
            // 
            this.panel36.Controls.Add(this.tbTenTaiKhoan);
            this.panel36.Controls.Add(this.label11);
            this.panel36.Location = new System.Drawing.Point(3, 3);
            this.panel36.Name = "panel36";
            this.panel36.Size = new System.Drawing.Size(483, 56);
            this.panel36.TabIndex = 1;
            // 
            // tbTenTaiKhoan
            // 
            this.tbTenTaiKhoan.Location = new System.Drawing.Point(209, 14);
            this.tbTenTaiKhoan.Name = "tbTenTaiKhoan";
            this.tbTenTaiKhoan.Size = new System.Drawing.Size(254, 22);
            this.tbTenTaiKhoan.TabIndex = 1;
            this.tbTenTaiKhoan.TextChanged += new System.EventHandler(this.tbTenTaiKhoan_TextChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label11.Location = new System.Drawing.Point(3, 14);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(178, 29);
            this.label11.TabIndex = 0;
            this.label11.Text = "Tên tài khoản:";
            // 
            // bCancel
            // 
            this.bCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.bCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.bCancel.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.bCancel.Location = new System.Drawing.Point(318, 361);
            this.bCancel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.bCancel.Name = "bCancel";
            this.bCancel.Size = new System.Drawing.Size(180, 44);
            this.bCancel.TabIndex = 10;
            this.bCancel.Text = "Cancel";
            this.bCancel.UseVisualStyleBackColor = false;
            this.bCancel.Click += new System.EventHandler(this.bCancel_Click);
            // 
            // bOK
            // 
            this.bOK.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.bOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.bOK.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.bOK.Location = new System.Drawing.Point(15, 361);
            this.bOK.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.bOK.Name = "bOK";
            this.bOK.Size = new System.Drawing.Size(174, 44);
            this.bOK.TabIndex = 9;
            this.bOK.Text = "OK";
            this.bOK.UseVisualStyleBackColor = false;
            this.bOK.Click += new System.EventHandler(this.bOK_Click);
            // 
            // AddAccount
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LemonChiffon;
            this.ClientSize = new System.Drawing.Size(513, 426);
            this.Controls.Add(this.bCancel);
            this.Controls.Add(this.bOK);
            this.Controls.Add(this.panel29);
            this.Name = "AddAccount";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AddAccount";
            this.panel29.ResumeLayout(false);
            this.panel46.ResumeLayout(false);
            this.panel46.PerformLayout();
            this.panel38.ResumeLayout(false);
            this.panel38.PerformLayout();
            this.panel34.ResumeLayout(false);
            this.panel34.PerformLayout();
            this.panel35.ResumeLayout(false);
            this.panel35.PerformLayout();
            this.panel36.ResumeLayout(false);
            this.panel36.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel29;
        private System.Windows.Forms.Panel panel46;
        private System.Windows.Forms.ComboBox cbMaNhanVien;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Panel panel38;
        private System.Windows.Forms.ComboBox cbLoaiTaiKhoan;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Panel panel34;
        private System.Windows.Forms.TextBox tbMatKhau;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel35;
        private System.Windows.Forms.TextBox tbTenHienThi;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Panel panel36;
        private System.Windows.Forms.TextBox tbTenTaiKhoan;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button bCancel;
        private System.Windows.Forms.Button bOK;
    }
}