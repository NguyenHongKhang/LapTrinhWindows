
namespace WindowsFormsApp1
{
    partial class Payments2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Payments2));
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel9 = new System.Windows.Forms.Panel();
            this.tbSauGiamGia = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.tbSoTienKhachTra = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.tbSoTienThoiLai = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.panel5 = new System.Windows.Forms.Panel();
            this.cbGiamGia = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.panel6 = new System.Windows.Forms.Panel();
            this.tbTongTien = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btThanhToan = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.panel9.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel5.SuspendLayout();
            this.panel6.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel9);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.panel5);
            this.panel1.Controls.Add(this.panel6);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(491, 311);
            this.panel1.TabIndex = 8;
            // 
            // panel9
            // 
            this.panel9.Controls.Add(this.tbSauGiamGia);
            this.panel9.Controls.Add(this.label6);
            this.panel9.Location = new System.Drawing.Point(5, 127);
            this.panel9.Name = "panel9";
            this.panel9.Size = new System.Drawing.Size(483, 56);
            this.panel9.TabIndex = 6;
            // 
            // tbSauGiamGia
            // 
            this.tbSauGiamGia.Location = new System.Drawing.Point(226, 14);
            this.tbSauGiamGia.Name = "tbSauGiamGia";
            this.tbSauGiamGia.ReadOnly = true;
            this.tbSauGiamGia.Size = new System.Drawing.Size(254, 22);
            this.tbSauGiamGia.TabIndex = 2;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label6.Location = new System.Drawing.Point(3, 14);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(131, 29);
            this.label6.TabIndex = 0;
            this.label6.Text = "Tổng tiền:";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tbSoTienKhachTra);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Location = new System.Drawing.Point(5, 189);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(483, 56);
            this.panel2.TabIndex = 5;
            // 
            // tbSoTienKhachTra
            // 
            this.tbSoTienKhachTra.Location = new System.Drawing.Point(226, 14);
            this.tbSoTienKhachTra.Name = "tbSoTienKhachTra";
            this.tbSoTienKhachTra.Size = new System.Drawing.Size(254, 22);
            this.tbSoTienKhachTra.TabIndex = 2;
            this.tbSoTienKhachTra.Leave += new System.EventHandler(this.tbSoTienKhachTra_Leave);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label2.Location = new System.Drawing.Point(3, 14);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(214, 29);
            this.label2.TabIndex = 0;
            this.label2.Text = "Số tiền khách trả:";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.tbSoTienThoiLai);
            this.panel3.Controls.Add(this.label3);
            this.panel3.Location = new System.Drawing.Point(5, 250);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(483, 56);
            this.panel3.TabIndex = 4;
            // 
            // tbSoTienThoiLai
            // 
            this.tbSoTienThoiLai.Location = new System.Drawing.Point(226, 14);
            this.tbSoTienThoiLai.Name = "tbSoTienThoiLai";
            this.tbSoTienThoiLai.ReadOnly = true;
            this.tbSoTienThoiLai.Size = new System.Drawing.Size(254, 22);
            this.tbSoTienThoiLai.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label3.Location = new System.Drawing.Point(3, 14);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(187, 29);
            this.label3.TabIndex = 0;
            this.label3.Text = "Số tiền thối lại:";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.cbGiamGia);
            this.panel5.Controls.Add(this.label4);
            this.panel5.Location = new System.Drawing.Point(5, 65);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(483, 56);
            this.panel5.TabIndex = 3;
            // 
            // cbGiamGia
            // 
            this.cbGiamGia.FormattingEnabled = true;
            this.cbGiamGia.Location = new System.Drawing.Point(226, 14);
            this.cbGiamGia.Name = "cbGiamGia";
            this.cbGiamGia.Size = new System.Drawing.Size(254, 24);
            this.cbGiamGia.TabIndex = 2;
            this.cbGiamGia.SelectedIndexChanged += new System.EventHandler(this.cbGiamGia_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label4.Location = new System.Drawing.Point(3, 14);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(124, 29);
            this.label4.TabIndex = 0;
            this.label4.Text = "Giảm giá:";
            // 
            // panel6
            // 
            this.panel6.Controls.Add(this.tbTongTien);
            this.panel6.Controls.Add(this.label5);
            this.panel6.Location = new System.Drawing.Point(5, 3);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(483, 56);
            this.panel6.TabIndex = 2;
            // 
            // tbTongTien
            // 
            this.tbTongTien.Location = new System.Drawing.Point(226, 14);
            this.tbTongTien.Name = "tbTongTien";
            this.tbTongTien.ReadOnly = true;
            this.tbTongTien.Size = new System.Drawing.Size(254, 22);
            this.tbTongTien.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label5.Location = new System.Drawing.Point(3, 14);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(131, 29);
            this.label5.TabIndex = 0;
            this.label5.Text = "Tạm Tính:";
            // 
            // btThanhToan
            // 
            this.btThanhToan.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.btThanhToan.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btThanhToan.BackgroundImage")));
            this.btThanhToan.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btThanhToan.Location = new System.Drawing.Point(350, 329);
            this.btThanhToan.Name = "btThanhToan";
            this.btThanhToan.Size = new System.Drawing.Size(150, 51);
            this.btThanhToan.TabIndex = 7;
            this.btThanhToan.UseVisualStyleBackColor = false;
            this.btThanhToan.Click += new System.EventHandler(this.btThanhToan_Click);
            // 
            // Payments2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LemonChiffon;
            this.ClientSize = new System.Drawing.Size(511, 396);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btThanhToan);
            this.Name = "Payments2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Payments2";
            this.panel1.ResumeLayout(false);
            this.panel9.ResumeLayout(false);
            this.panel9.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel9;
        private System.Windows.Forms.TextBox tbSauGiamGia;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox tbSoTienKhachTra;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TextBox tbSoTienThoiLai;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.ComboBox cbGiamGia;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.TextBox tbTongTien;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btThanhToan;
    }
}