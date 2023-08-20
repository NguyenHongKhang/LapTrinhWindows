
namespace WindowsFormsApp1.Views
{
    partial class AddCustomer
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
            this.panel8 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tbSDT = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.bCancel = new System.Windows.Forms.Button();
            this.bOK = new System.Windows.Forms.Button();
            this.panel10 = new System.Windows.Forms.Panel();
            this.tbTenKhachHang = new System.Windows.Forms.TextBox();
            this.lbFoodName = new System.Windows.Forms.Label();
            this.panel9 = new System.Windows.Forms.Panel();
            this.IDKhachHang = new System.Windows.Forms.TextBox();
            this.lbIDFood = new System.Windows.Forms.Label();
            this.panel8.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel10.SuspendLayout();
            this.panel9.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel8
            // 
            this.panel8.BackColor = System.Drawing.Color.LemonChiffon;
            this.panel8.Controls.Add(this.panel1);
            this.panel8.Controls.Add(this.bCancel);
            this.panel8.Controls.Add(this.bOK);
            this.panel8.Controls.Add(this.panel10);
            this.panel8.Controls.Add(this.panel9);
            this.panel8.Location = new System.Drawing.Point(12, 12);
            this.panel8.Name = "panel8";
            this.panel8.Size = new System.Drawing.Size(517, 247);
            this.panel8.TabIndex = 6;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.tbSDT);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(3, 127);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(499, 56);
            this.panel1.TabIndex = 11;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // tbSDT
            // 
            this.tbSDT.Location = new System.Drawing.Point(226, 14);
            this.tbSDT.Name = "tbSDT";
            this.tbSDT.Size = new System.Drawing.Size(254, 22);
            this.tbSDT.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label1.Location = new System.Drawing.Point(3, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(174, 29);
            this.label1.TabIndex = 0;
            this.label1.Text = "Số điện thoại:";
            // 
            // bCancel
            // 
            this.bCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.bCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.bCancel.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.bCancel.Location = new System.Drawing.Point(306, 188);
            this.bCancel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.bCancel.Name = "bCancel";
            this.bCancel.Size = new System.Drawing.Size(180, 44);
            this.bCancel.TabIndex = 8;
            this.bCancel.Text = "Cancel";
            this.bCancel.UseVisualStyleBackColor = false;
            this.bCancel.Click += new System.EventHandler(this.bCancel_Click);
            // 
            // bOK
            // 
            this.bOK.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.bOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.bOK.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.bOK.Location = new System.Drawing.Point(0, 188);
            this.bOK.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.bOK.Name = "bOK";
            this.bOK.Size = new System.Drawing.Size(174, 44);
            this.bOK.TabIndex = 6;
            this.bOK.Text = "OK";
            this.bOK.UseVisualStyleBackColor = false;
            this.bOK.Click += new System.EventHandler(this.bOK_Click);
            // 
            // panel10
            // 
            this.panel10.Controls.Add(this.tbTenKhachHang);
            this.panel10.Controls.Add(this.lbFoodName);
            this.panel10.Location = new System.Drawing.Point(3, 65);
            this.panel10.Name = "panel10";
            this.panel10.Size = new System.Drawing.Size(499, 56);
            this.panel10.TabIndex = 2;
            // 
            // tbTenKhachHang
            // 
            this.tbTenKhachHang.Location = new System.Drawing.Point(229, 14);
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
            this.lbFoodName.Size = new System.Drawing.Size(213, 29);
            this.lbFoodName.TabIndex = 0;
            this.lbFoodName.Text = "Tên Khách Hàng:";
            // 
            // panel9
            // 
            this.panel9.Controls.Add(this.IDKhachHang);
            this.panel9.Controls.Add(this.lbIDFood);
            this.panel9.Location = new System.Drawing.Point(3, 3);
            this.panel9.Name = "panel9";
            this.panel9.Size = new System.Drawing.Size(499, 56);
            this.panel9.TabIndex = 1;
            // 
            // IDKhachHang
            // 
            this.IDKhachHang.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.IDKhachHang.Location = new System.Drawing.Point(229, 14);
            this.IDKhachHang.Name = "IDKhachHang";
            this.IDKhachHang.ReadOnly = true;
            this.IDKhachHang.Size = new System.Drawing.Size(254, 22);
            this.IDKhachHang.TabIndex = 1;
            // 
            // lbIDFood
            // 
            this.lbIDFood.AutoSize = true;
            this.lbIDFood.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lbIDFood.Location = new System.Drawing.Point(3, 14);
            this.lbIDFood.Name = "lbIDFood";
            this.lbIDFood.Size = new System.Drawing.Size(202, 29);
            this.lbIDFood.TabIndex = 0;
            this.lbIDFood.Text = "Mã Khách Hàng:";
            // 
            // AddCustomer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LemonChiffon;
            this.ClientSize = new System.Drawing.Size(533, 266);
            this.Controls.Add(this.panel8);
            this.Name = "AddCustomer";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AddCustomer";
            this.panel8.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel10.ResumeLayout(false);
            this.panel10.PerformLayout();
            this.panel9.ResumeLayout(false);
            this.panel9.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel8;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox tbSDT;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button bCancel;
        private System.Windows.Forms.Button bOK;
        private System.Windows.Forms.Panel panel10;
        private System.Windows.Forms.TextBox tbTenKhachHang;
        private System.Windows.Forms.Label lbFoodName;
        private System.Windows.Forms.Panel panel9;
        private System.Windows.Forms.TextBox IDKhachHang;
        private System.Windows.Forms.Label lbIDFood;
    }
}