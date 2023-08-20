
namespace WindowsFormsApp1
{
    partial class Customer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Customer));
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel7 = new System.Windows.Forms.Panel();
            this.tbfind = new System.Windows.Forms.TextBox();
            this.btfind = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.dgvCustomer = new System.Windows.Forms.DataGridView();
            this.panel8 = new System.Windows.Forms.Panel();
            this.panel15 = new System.Windows.Forms.Panel();
            this.tbSoBillsDaThanhToan = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel12 = new System.Windows.Forms.Panel();
            this.tbDiemTichLuy = new System.Windows.Forms.TextBox();
            this.lbPrice = new System.Windows.Forms.Label();
            this.panel11 = new System.Windows.Forms.Panel();
            this.tbSoDienThoai = new System.Windows.Forms.TextBox();
            this.lbCategory = new System.Windows.Forms.Label();
            this.panel10 = new System.Windows.Forms.Panel();
            this.tbTenKhachHang = new System.Windows.Forms.TextBox();
            this.lbFoodName = new System.Windows.Forms.Label();
            this.panel9 = new System.Windows.Forms.Panel();
            this.tbMaKhachHang = new System.Windows.Forms.TextBox();
            this.lbIDFood = new System.Windows.Forms.Label();
            this.panel5 = new System.Windows.Forms.Panel();
            this.buttonDeleteCustomer = new System.Windows.Forms.Button();
            this.buttonFixCustomer = new System.Windows.Forms.Button();
            this.buttonAddCustomer = new System.Windows.Forms.Button();
            this.panel4.SuspendLayout();
            this.panel7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCustomer)).BeginInit();
            this.panel8.SuspendLayout();
            this.panel15.SuspendLayout();
            this.panel12.SuspendLayout();
            this.panel11.SuspendLayout();
            this.panel10.SuspendLayout();
            this.panel9.SuspendLayout();
            this.panel5.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.panel7);
            this.panel4.Controls.Add(this.pictureBox1);
            this.panel4.Controls.Add(this.dgvCustomer);
            this.panel4.Controls.Add(this.panel8);
            this.panel4.Controls.Add(this.panel5);
            this.panel4.Location = new System.Drawing.Point(0, 0);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1320, 710);
            this.panel4.TabIndex = 1;
            this.panel4.Paint += new System.Windows.Forms.PaintEventHandler(this.panel4_Paint);
            // 
            // panel7
            // 
            this.panel7.Controls.Add(this.tbfind);
            this.panel7.Controls.Add(this.btfind);
            this.panel7.Location = new System.Drawing.Point(1046, 302);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(263, 41);
            this.panel7.TabIndex = 13;
            // 
            // tbfind
            // 
            this.tbfind.Location = new System.Drawing.Point(6, 11);
            this.tbfind.Name = "tbfind";
            this.tbfind.Size = new System.Drawing.Size(159, 22);
            this.tbfind.TabIndex = 3;
            // 
            // btfind
            // 
            this.btfind.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.btfind.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btfind.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btfind.Location = new System.Drawing.Point(171, 10);
            this.btfind.Name = "btfind";
            this.btfind.Size = new System.Drawing.Size(89, 24);
            this.btfind.TabIndex = 2;
            this.btfind.Text = "Tìm kiếm";
            this.btfind.UseVisualStyleBackColor = false;
            this.btfind.Click += new System.EventHandler(this.btfind_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(515, 22);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(517, 321);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            // 
            // dgvCustomer
            // 
            this.dgvCustomer.AllowUserToAddRows = false;
            this.dgvCustomer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvCustomer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCustomer.Location = new System.Drawing.Point(6, 353);
            this.dgvCustomer.Name = "dgvCustomer";
            this.dgvCustomer.ReadOnly = true;
            this.dgvCustomer.RowHeadersWidth = 51;
            this.dgvCustomer.RowTemplate.Height = 24;
            this.dgvCustomer.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvCustomer.Size = new System.Drawing.Size(1303, 345);
            this.dgvCustomer.TabIndex = 0;
            // 
            // panel8
            // 
            this.panel8.Controls.Add(this.panel15);
            this.panel8.Controls.Add(this.panel12);
            this.panel8.Controls.Add(this.panel11);
            this.panel8.Controls.Add(this.panel10);
            this.panel8.Controls.Add(this.panel9);
            this.panel8.Location = new System.Drawing.Point(6, 19);
            this.panel8.Name = "panel8";
            this.panel8.Size = new System.Drawing.Size(493, 324);
            this.panel8.TabIndex = 3;
            // 
            // panel15
            // 
            this.panel15.Controls.Add(this.tbSoBillsDaThanhToan);
            this.panel15.Controls.Add(this.label1);
            this.panel15.Location = new System.Drawing.Point(3, 189);
            this.panel15.Name = "panel15";
            this.panel15.Size = new System.Drawing.Size(483, 56);
            this.panel15.TabIndex = 5;
            // 
            // tbSoBillsDaThanhToan
            // 
            this.tbSoBillsDaThanhToan.Location = new System.Drawing.Point(226, 14);
            this.tbSoBillsDaThanhToan.Name = "tbSoBillsDaThanhToan";
            this.tbSoBillsDaThanhToan.ReadOnly = true;
            this.tbSoBillsDaThanhToan.Size = new System.Drawing.Size(254, 22);
            this.tbSoBillsDaThanhToan.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label1.Location = new System.Drawing.Point(3, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(117, 29);
            this.label1.TabIndex = 0;
            this.label1.Text = "Số Bills :";
            // 
            // panel12
            // 
            this.panel12.Controls.Add(this.tbDiemTichLuy);
            this.panel12.Controls.Add(this.lbPrice);
            this.panel12.Location = new System.Drawing.Point(3, 250);
            this.panel12.Name = "panel12";
            this.panel12.Size = new System.Drawing.Size(483, 56);
            this.panel12.TabIndex = 4;
            // 
            // tbDiemTichLuy
            // 
            this.tbDiemTichLuy.Location = new System.Drawing.Point(226, 14);
            this.tbDiemTichLuy.Name = "tbDiemTichLuy";
            this.tbDiemTichLuy.Size = new System.Drawing.Size(254, 22);
            this.tbDiemTichLuy.TabIndex = 3;
            // 
            // lbPrice
            // 
            this.lbPrice.AutoSize = true;
            this.lbPrice.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lbPrice.Location = new System.Drawing.Point(3, 14);
            this.lbPrice.Name = "lbPrice";
            this.lbPrice.Size = new System.Drawing.Size(81, 29);
            this.lbPrice.TabIndex = 0;
            this.lbPrice.Text = "Điểm:";
            // 
            // panel11
            // 
            this.panel11.Controls.Add(this.tbSoDienThoai);
            this.panel11.Controls.Add(this.lbCategory);
            this.panel11.Location = new System.Drawing.Point(3, 127);
            this.panel11.Name = "panel11";
            this.panel11.Size = new System.Drawing.Size(483, 56);
            this.panel11.TabIndex = 3;
            // 
            // tbSoDienThoai
            // 
            this.tbSoDienThoai.Location = new System.Drawing.Point(226, 14);
            this.tbSoDienThoai.Name = "tbSoDienThoai";
            this.tbSoDienThoai.Size = new System.Drawing.Size(254, 22);
            this.tbSoDienThoai.TabIndex = 2;
            // 
            // lbCategory
            // 
            this.lbCategory.AutoSize = true;
            this.lbCategory.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lbCategory.Location = new System.Drawing.Point(3, 14);
            this.lbCategory.Name = "lbCategory";
            this.lbCategory.Size = new System.Drawing.Size(174, 29);
            this.lbCategory.TabIndex = 0;
            this.lbCategory.Text = "Số điện thoại:";
            // 
            // panel10
            // 
            this.panel10.Controls.Add(this.tbTenKhachHang);
            this.panel10.Controls.Add(this.lbFoodName);
            this.panel10.Location = new System.Drawing.Point(3, 65);
            this.panel10.Name = "panel10";
            this.panel10.Size = new System.Drawing.Size(483, 56);
            this.panel10.TabIndex = 2;
            // 
            // tbTenKhachHang
            // 
            this.tbTenKhachHang.Location = new System.Drawing.Point(226, 14);
            this.tbTenKhachHang.Name = "tbTenKhachHang";
            this.tbTenKhachHang.Size = new System.Drawing.Size(254, 22);
            this.tbTenKhachHang.TabIndex = 1;
            // 
            // lbFoodName
            // 
            this.lbFoodName.AutoSize = true;
            this.lbFoodName.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lbFoodName.Location = new System.Drawing.Point(3, 14);
            this.lbFoodName.Name = "lbFoodName";
            this.lbFoodName.Size = new System.Drawing.Size(205, 29);
            this.lbFoodName.TabIndex = 0;
            this.lbFoodName.Text = "Tên khách hàng:";
            // 
            // panel9
            // 
            this.panel9.Controls.Add(this.tbMaKhachHang);
            this.panel9.Controls.Add(this.lbIDFood);
            this.panel9.Location = new System.Drawing.Point(3, 3);
            this.panel9.Name = "panel9";
            this.panel9.Size = new System.Drawing.Size(483, 56);
            this.panel9.TabIndex = 1;
            // 
            // tbMaKhachHang
            // 
            this.tbMaKhachHang.Location = new System.Drawing.Point(226, 18);
            this.tbMaKhachHang.Name = "tbMaKhachHang";
            this.tbMaKhachHang.ReadOnly = true;
            this.tbMaKhachHang.Size = new System.Drawing.Size(254, 22);
            this.tbMaKhachHang.TabIndex = 1;
            this.tbMaKhachHang.TextChanged += new System.EventHandler(this.tbMaKhachHang_TextChanged);
            // 
            // lbIDFood
            // 
            this.lbIDFood.AutoSize = true;
            this.lbIDFood.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lbIDFood.Location = new System.Drawing.Point(3, 14);
            this.lbIDFood.Name = "lbIDFood";
            this.lbIDFood.Size = new System.Drawing.Size(194, 29);
            this.lbIDFood.TabIndex = 0;
            this.lbIDFood.Text = "Mã khách hàng:";
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.buttonDeleteCustomer);
            this.panel5.Controls.Add(this.buttonFixCustomer);
            this.panel5.Controls.Add(this.buttonAddCustomer);
            this.panel5.Location = new System.Drawing.Point(1049, 22);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(263, 180);
            this.panel5.TabIndex = 0;
            // 
            // buttonDeleteCustomer
            // 
            this.buttonDeleteCustomer.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.buttonDeleteCustomer.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.buttonDeleteCustomer.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.buttonDeleteCustomer.Location = new System.Drawing.Point(23, 122);
            this.buttonDeleteCustomer.Name = "buttonDeleteCustomer";
            this.buttonDeleteCustomer.Size = new System.Drawing.Size(230, 51);
            this.buttonDeleteCustomer.TabIndex = 2;
            this.buttonDeleteCustomer.Text = "Xóa";
            this.buttonDeleteCustomer.UseVisualStyleBackColor = false;
            this.buttonDeleteCustomer.Click += new System.EventHandler(this.buttonDeleteCustomer_Click);
            // 
            // buttonFixCustomer
            // 
            this.buttonFixCustomer.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.buttonFixCustomer.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.buttonFixCustomer.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.buttonFixCustomer.Location = new System.Drawing.Point(23, 65);
            this.buttonFixCustomer.Name = "buttonFixCustomer";
            this.buttonFixCustomer.Size = new System.Drawing.Size(230, 51);
            this.buttonFixCustomer.TabIndex = 1;
            this.buttonFixCustomer.Text = "Sửa";
            this.buttonFixCustomer.UseVisualStyleBackColor = false;
            this.buttonFixCustomer.Click += new System.EventHandler(this.buttonFixCustomer_Click);
            // 
            // buttonAddCustomer
            // 
            this.buttonAddCustomer.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.buttonAddCustomer.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.buttonAddCustomer.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.buttonAddCustomer.Location = new System.Drawing.Point(23, 0);
            this.buttonAddCustomer.Name = "buttonAddCustomer";
            this.buttonAddCustomer.Size = new System.Drawing.Size(230, 59);
            this.buttonAddCustomer.TabIndex = 0;
            this.buttonAddCustomer.Text = "Thêm";
            this.buttonAddCustomer.UseVisualStyleBackColor = false;
            this.buttonAddCustomer.Click += new System.EventHandler(this.buttonAddCustomer_Click);
            // 
            // Customer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LemonChiffon;
            this.ClientSize = new System.Drawing.Size(1327, 720);
            this.Controls.Add(this.panel4);
            this.Name = "Customer";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Thông tin khách hàng";
            this.panel4.ResumeLayout(false);
            this.panel7.ResumeLayout(false);
            this.panel7.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCustomer)).EndInit();
            this.panel8.ResumeLayout(false);
            this.panel15.ResumeLayout(false);
            this.panel15.PerformLayout();
            this.panel12.ResumeLayout(false);
            this.panel12.PerformLayout();
            this.panel11.ResumeLayout(false);
            this.panel11.PerformLayout();
            this.panel10.ResumeLayout(false);
            this.panel10.PerformLayout();
            this.panel9.ResumeLayout(false);
            this.panel9.PerformLayout();
            this.panel5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.DataGridView dgvCustomer;
        private System.Windows.Forms.Panel panel8;
        private System.Windows.Forms.Panel panel15;
        private System.Windows.Forms.TextBox tbSoBillsDaThanhToan;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel12;
        private System.Windows.Forms.TextBox tbDiemTichLuy;
        private System.Windows.Forms.Label lbPrice;
        private System.Windows.Forms.Panel panel11;
        private System.Windows.Forms.TextBox tbSoDienThoai;
        private System.Windows.Forms.Label lbCategory;
        private System.Windows.Forms.Panel panel10;
        private System.Windows.Forms.TextBox tbTenKhachHang;
        private System.Windows.Forms.Label lbFoodName;
        private System.Windows.Forms.Panel panel9;
        private System.Windows.Forms.TextBox tbMaKhachHang;
        private System.Windows.Forms.Label lbIDFood;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Button buttonDeleteCustomer;
        private System.Windows.Forms.Button buttonFixCustomer;
        private System.Windows.Forms.Button buttonAddCustomer;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.TextBox tbfind;
        private System.Windows.Forms.Button btfind;
    }
}