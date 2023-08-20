
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    partial class TableManagement
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TableManagement));
            this.lsvBill = new System.Windows.Forms.ListView();
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.flpTable = new System.Windows.Forms.FlowLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonAdd = new System.Windows.Forms.Button();
            this.tbTongTien = new System.Windows.Forms.TextBox();
            this.buttonPay = new System.Windows.Forms.Button();
            this.cbMoveTable = new System.Windows.Forms.ComboBox();
            this.btMoveTable = new System.Windows.Forms.Button();
            this.btCancelTable = new System.Windows.Forms.Button();
            this.btTableMerge = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.adminToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiQuanly = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiThongke = new System.Windows.Forms.ToolStripMenuItem();
            this.thôngTinTàiKhoảnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiThongtin = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiThaydoi = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiDangxuat = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiKhachhang = new System.Windows.Forms.ToolStripMenuItem();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lsvBill
            // 
            this.lsvBill.Alignment = System.Windows.Forms.ListViewAlignment.SnapToGrid;
            this.lsvBill.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader3,
            this.columnHeader4,
            this.columnHeader5,
            this.columnHeader1});
            this.lsvBill.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lsvBill.GridLines = true;
            this.lsvBill.HideSelection = false;
            this.lsvBill.Location = new System.Drawing.Point(579, 28);
            this.lsvBill.Name = "lsvBill";
            this.lsvBill.Scrollable = false;
            this.lsvBill.Size = new System.Drawing.Size(733, 725);
            this.lsvBill.TabIndex = 0;
            this.lsvBill.UseCompatibleStateImageBehavior = false;
            this.lsvBill.View = System.Windows.Forms.View.Details;
            this.lsvBill.SelectedIndexChanged += new System.EventHandler(this.lsvBill_SelectedIndexChanged);
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Tên Món";
            this.columnHeader3.Width = 150;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Số Lượng";
            this.columnHeader4.Width = 100;
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "Đơn Giá";
            this.columnHeader5.Width = 150;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Thành Tiền";
            this.columnHeader1.Width = 500;
            // 
            // flpTable
            // 
            this.flpTable.AutoScroll = true;
            this.flpTable.BackColor = System.Drawing.Color.LemonChiffon;
            this.flpTable.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.flpTable.Dock = System.Windows.Forms.DockStyle.Left;
            this.flpTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.flpTable.Location = new System.Drawing.Point(0, 28);
            this.flpTable.Name = "flpTable";
            this.flpTable.Size = new System.Drawing.Size(579, 725);
            this.flpTable.TabIndex = 5;
            this.flpTable.Paint += new System.Windows.Forms.PaintEventHandler(this.flpTable_Paint);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(54, 626);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(142, 32);
            this.label2.TabIndex = 9;
            this.label2.Text = "Tổng tiền";
            // 
            // buttonAdd
            // 
            this.buttonAdd.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.buttonAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.buttonAdd.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.buttonAdd.Location = new System.Drawing.Point(19, 244);
            this.buttonAdd.Name = "buttonAdd";
            this.buttonAdd.Size = new System.Drawing.Size(217, 78);
            this.buttonAdd.TabIndex = 2;
            this.buttonAdd.Text = "Thêm Món";
            this.buttonAdd.UseVisualStyleBackColor = false;
            this.buttonAdd.Click += new System.EventHandler(this.buttonAdd_Click);
            // 
            // tbTongTien
            // 
            this.tbTongTien.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.tbTongTien.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.tbTongTien.ForeColor = System.Drawing.Color.Red;
            this.tbTongTien.Location = new System.Drawing.Point(19, 661);
            this.tbTongTien.Name = "tbTongTien";
            this.tbTongTien.ReadOnly = true;
            this.tbTongTien.Size = new System.Drawing.Size(216, 34);
            this.tbTongTien.TabIndex = 12;
            this.tbTongTien.Text = "0";
            this.tbTongTien.TextChanged += new System.EventHandler(this.tbTongTien_TextChanged);
            // 
            // buttonPay
            // 
            this.buttonPay.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.buttonPay.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.buttonPay.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.buttonPay.Location = new System.Drawing.Point(19, 539);
            this.buttonPay.Name = "buttonPay";
            this.buttonPay.Size = new System.Drawing.Size(216, 78);
            this.buttonPay.TabIndex = 3;
            this.buttonPay.Text = "Thanh toán";
            this.buttonPay.UseVisualStyleBackColor = false;
            this.buttonPay.Click += new System.EventHandler(this.buttonPay_Click);
            // 
            // cbMoveTable
            // 
            this.cbMoveTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.cbMoveTable.FormattingEnabled = true;
            this.cbMoveTable.Location = new System.Drawing.Point(19, 328);
            this.cbMoveTable.Name = "cbMoveTable";
            this.cbMoveTable.Size = new System.Drawing.Size(217, 37);
            this.cbMoveTable.TabIndex = 7;
            // 
            // btMoveTable
            // 
            this.btMoveTable.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.btMoveTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btMoveTable.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btMoveTable.Location = new System.Drawing.Point(19, 371);
            this.btMoveTable.Name = "btMoveTable";
            this.btMoveTable.Size = new System.Drawing.Size(216, 78);
            this.btMoveTable.TabIndex = 6;
            this.btMoveTable.Text = "Chuyển bàn";
            this.btMoveTable.UseVisualStyleBackColor = false;
            this.btMoveTable.Click += new System.EventHandler(this.btMoveTable_Click);
            // 
            // btCancelTable
            // 
            this.btCancelTable.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.btCancelTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btCancelTable.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btCancelTable.Location = new System.Drawing.Point(19, 160);
            this.btCancelTable.Name = "btCancelTable";
            this.btCancelTable.Size = new System.Drawing.Size(217, 78);
            this.btCancelTable.TabIndex = 13;
            this.btCancelTable.Text = "Hủy bàn";
            this.btCancelTable.UseVisualStyleBackColor = false;
            this.btCancelTable.Click += new System.EventHandler(this.btCancelTable_Click);
            // 
            // btTableMerge
            // 
            this.btTableMerge.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.btTableMerge.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btTableMerge.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btTableMerge.Location = new System.Drawing.Point(19, 455);
            this.btTableMerge.Name = "btTableMerge";
            this.btTableMerge.Size = new System.Drawing.Size(216, 78);
            this.btTableMerge.TabIndex = 14;
            this.btTableMerge.Text = "Gộp bàn";
            this.btTableMerge.UseCompatibleTextRendering = true;
            this.btTableMerge.UseVisualStyleBackColor = false;
            this.btTableMerge.Click += new System.EventHandler(this.btTableMerge_Click);
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.LemonChiffon;
            this.panel3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel3.Controls.Add(this.btTableMerge);
            this.panel3.Controls.Add(this.btCancelTable);
            this.panel3.Controls.Add(this.btMoveTable);
            this.panel3.Controls.Add(this.cbMoveTable);
            this.panel3.Controls.Add(this.buttonPay);
            this.panel3.Controls.Add(this.tbTongTien);
            this.panel3.Controls.Add(this.buttonAdd);
            this.panel3.Controls.Add(this.pictureBox1);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel3.Location = new System.Drawing.Point(1312, 28);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(270, 725);
            this.panel3.TabIndex = 3;
            this.panel3.Paint += new System.Windows.Forms.PaintEventHandler(this.panel3_Paint);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::WindowsFormsApp1.Properties.Resources.cute_coffee_shop_building_icons_elements_with_flat_design_7130_1336;
            this.pictureBox1.Location = new System.Drawing.Point(19, 13);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(216, 141);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 11;
            this.pictureBox1.TabStop = false;
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("menuStrip1.BackgroundImage")));
            this.menuStrip1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.adminToolStripMenuItem,
            this.thôngTinTàiKhoảnToolStripMenuItem,
            this.tsmiKhachhang});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1582, 28);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // adminToolStripMenuItem
            // 
            this.adminToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiQuanly,
            this.tsmiThongke});
            this.adminToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("adminToolStripMenuItem.Image")));
            this.adminToolStripMenuItem.Name = "adminToolStripMenuItem";
            this.adminToolStripMenuItem.Size = new System.Drawing.Size(87, 24);
            this.adminToolStripMenuItem.Text = "Admin";
            // 
            // tsmiQuanly
            // 
            this.tsmiQuanly.Image = ((System.Drawing.Image)(resources.GetObject("tsmiQuanly.Image")));
            this.tsmiQuanly.Name = "tsmiQuanly";
            this.tsmiQuanly.Size = new System.Drawing.Size(153, 26);
            this.tsmiQuanly.Text = "Quản lý ";
            this.tsmiQuanly.Click += new System.EventHandler(this.tsmiQuanly_Click);
            // 
            // tsmiThongke
            // 
            this.tsmiThongke.Image = ((System.Drawing.Image)(resources.GetObject("tsmiThongke.Image")));
            this.tsmiThongke.Name = "tsmiThongke";
            this.tsmiThongke.Size = new System.Drawing.Size(153, 26);
            this.tsmiThongke.Text = "Thống kê";
            this.tsmiThongke.Click += new System.EventHandler(this.tsmiThongke_Click);
            // 
            // thôngTinTàiKhoảnToolStripMenuItem
            // 
            this.thôngTinTàiKhoảnToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiThongtin,
            this.tsmiThaydoi,
            this.tsmiDangxuat});
            this.thôngTinTàiKhoảnToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("thôngTinTàiKhoảnToolStripMenuItem.Image")));
            this.thôngTinTàiKhoảnToolStripMenuItem.Name = "thôngTinTàiKhoảnToolStripMenuItem";
            this.thôngTinTàiKhoảnToolStripMenuItem.Size = new System.Drawing.Size(171, 24);
            this.thôngTinTàiKhoảnToolStripMenuItem.Text = "Thông tin tài khoản";
            // 
            // tsmiThongtin
            // 
            this.tsmiThongtin.Image = ((System.Drawing.Image)(resources.GetObject("tsmiThongtin.Image")));
            this.tsmiThongtin.Name = "tsmiThongtin";
            this.tsmiThongtin.Size = new System.Drawing.Size(218, 26);
            this.tsmiThongtin.Text = "Thông tin cá nhân";
            this.tsmiThongtin.Click += new System.EventHandler(this.tsmiThongtin_Click);
            // 
            // tsmiThaydoi
            // 
            this.tsmiThaydoi.Image = ((System.Drawing.Image)(resources.GetObject("tsmiThaydoi.Image")));
            this.tsmiThaydoi.Name = "tsmiThaydoi";
            this.tsmiThaydoi.Size = new System.Drawing.Size(218, 26);
            this.tsmiThaydoi.Text = "Thay đổi tài khoản ";
            this.tsmiThaydoi.Click += new System.EventHandler(this.tsmiThaydoi_Click);
            // 
            // tsmiDangxuat
            // 
            this.tsmiDangxuat.Image = ((System.Drawing.Image)(resources.GetObject("tsmiDangxuat.Image")));
            this.tsmiDangxuat.Name = "tsmiDangxuat";
            this.tsmiDangxuat.Size = new System.Drawing.Size(218, 26);
            this.tsmiDangxuat.Text = "Đăng xuất";
            this.tsmiDangxuat.Click += new System.EventHandler(this.tsmiDangxuat_Click);
            // 
            // tsmiKhachhang
            // 
            this.tsmiKhachhang.Image = ((System.Drawing.Image)(resources.GetObject("tsmiKhachhang.Image")));
            this.tsmiKhachhang.Name = "tsmiKhachhang";
            this.tsmiKhachhang.Size = new System.Drawing.Size(185, 24);
            this.tsmiKhachhang.Text = "Thông tin khách hàng";
            this.tsmiKhachhang.Click += new System.EventHandler(this.tsmiKhachhang_Click);
            // 
            // TableManagement
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1582, 753);
            this.Controls.Add(this.lsvBill);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.flpTable);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(1127, 610);
            this.Name = "TableManagement";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Caffee Shop";
            this.Load += new System.EventHandler(this.TableManagement_Load);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem adminToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem thôngTinTàiKhoảnToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tsmiThaydoi;
        private System.Windows.Forms.ToolStripMenuItem tsmiDangxuat;
        private System.Windows.Forms.ListView lsvBill;
        private System.Windows.Forms.FlowLayoutPanel flpTable;
        private System.Windows.Forms.ToolStripMenuItem tsmiKhachhang;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.ColumnHeader columnHeader5;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ToolStripMenuItem tsmiQuanly;
        private System.Windows.Forms.ToolStripMenuItem tsmiThongke;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button buttonAdd;
        private System.Windows.Forms.TextBox tbTongTien;
        private System.Windows.Forms.Button buttonPay;
        private System.Windows.Forms.ComboBox cbMoveTable;
        private System.Windows.Forms.Button btMoveTable;
        private System.Windows.Forms.Button btCancelTable;
        private System.Windows.Forms.Button btTableMerge;
        private System.Windows.Forms.Panel panel3;
        private ToolStripMenuItem tsmiThongtin;
    }
}