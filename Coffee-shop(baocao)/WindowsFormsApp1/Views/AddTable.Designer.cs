
namespace WindowsFormsApp1.Views
{
    partial class AddTable
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
            this.panel45 = new System.Windows.Forms.Panel();
            this.tbTrangThai = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.panel32 = new System.Windows.Forms.Panel();
            this.tbTenBan = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.panel33 = new System.Windows.Forms.Panel();
            this.tbMaBan = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.bCancel = new System.Windows.Forms.Button();
            this.bOK = new System.Windows.Forms.Button();
            this.panel45.SuspendLayout();
            this.panel32.SuspendLayout();
            this.panel33.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel45
            // 
            this.panel45.Controls.Add(this.tbTrangThai);
            this.panel45.Controls.Add(this.label10);
            this.panel45.Location = new System.Drawing.Point(12, 136);
            this.panel45.Name = "panel45";
            this.panel45.Size = new System.Drawing.Size(483, 56);
            this.panel45.TabIndex = 6;
            // 
            // tbTrangThai
            // 
            this.tbTrangThai.Location = new System.Drawing.Point(226, 12);
            this.tbTrangThai.Name = "tbTrangThai";
            this.tbTrangThai.ReadOnly = true;
            this.tbTrangThai.Size = new System.Drawing.Size(254, 22);
            this.tbTrangThai.TabIndex = 1;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label10.Location = new System.Drawing.Point(3, 14);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(138, 29);
            this.label10.TabIndex = 0;
            this.label10.Text = "Trạng thái:";
            // 
            // panel32
            // 
            this.panel32.Controls.Add(this.tbTenBan);
            this.panel32.Controls.Add(this.label8);
            this.panel32.Location = new System.Drawing.Point(12, 74);
            this.panel32.Name = "panel32";
            this.panel32.Size = new System.Drawing.Size(483, 56);
            this.panel32.TabIndex = 5;
            // 
            // tbTenBan
            // 
            this.tbTenBan.Location = new System.Drawing.Point(226, 14);
            this.tbTenBan.Name = "tbTenBan";
            this.tbTenBan.Size = new System.Drawing.Size(254, 22);
            this.tbTenBan.TabIndex = 1;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label8.Location = new System.Drawing.Point(3, 14);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(116, 29);
            this.label8.TabIndex = 0;
            this.label8.Text = "Tên bàn:";
            // 
            // panel33
            // 
            this.panel33.Controls.Add(this.tbMaBan);
            this.panel33.Controls.Add(this.label9);
            this.panel33.Location = new System.Drawing.Point(12, 12);
            this.panel33.Name = "panel33";
            this.panel33.Size = new System.Drawing.Size(483, 56);
            this.panel33.TabIndex = 4;
            // 
            // tbMaBan
            // 
            this.tbMaBan.Location = new System.Drawing.Point(226, 12);
            this.tbMaBan.Name = "tbMaBan";
            this.tbMaBan.ReadOnly = true;
            this.tbMaBan.Size = new System.Drawing.Size(254, 22);
            this.tbMaBan.TabIndex = 1;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label9.Location = new System.Drawing.Point(3, 14);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(105, 29);
            this.label9.TabIndex = 0;
            this.label9.Text = "Mã bàn:";
            // 
            // bCancel
            // 
            this.bCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.bCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.bCancel.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.bCancel.Location = new System.Drawing.Point(312, 216);
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
            this.bOK.Location = new System.Drawing.Point(12, 216);
            this.bOK.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.bOK.Name = "bOK";
            this.bOK.Size = new System.Drawing.Size(174, 44);
            this.bOK.TabIndex = 9;
            this.bOK.Text = "OK";
            this.bOK.UseVisualStyleBackColor = false;
            this.bOK.Click += new System.EventHandler(this.bOK_Click);
            // 
            // AddTable
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LemonChiffon;
            this.ClientSize = new System.Drawing.Size(504, 280);
            this.Controls.Add(this.bCancel);
            this.Controls.Add(this.bOK);
            this.Controls.Add(this.panel45);
            this.Controls.Add(this.panel32);
            this.Controls.Add(this.panel33);
            this.Name = "AddTable";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AddTable";
            this.panel45.ResumeLayout(false);
            this.panel45.PerformLayout();
            this.panel32.ResumeLayout(false);
            this.panel32.PerformLayout();
            this.panel33.ResumeLayout(false);
            this.panel33.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel45;
        private System.Windows.Forms.TextBox tbTrangThai;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Panel panel32;
        private System.Windows.Forms.TextBox tbTenBan;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Panel panel33;
        private System.Windows.Forms.TextBox tbMaBan;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button bCancel;
        private System.Windows.Forms.Button bOK;
    }
}