
namespace WindowsFormsApp1
{
    partial class AddFood
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
            this.label2 = new System.Windows.Forms.Label();
            this.bCancel = new System.Windows.Forms.Button();
            this.bOK = new System.Windows.Forms.Button();
            this.pbfood = new System.Windows.Forms.PictureBox();
            this.panel15 = new System.Windows.Forms.Panel();
            this.bFalse = new System.Windows.Forms.RadioButton();
            this.btrue = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.panel12 = new System.Windows.Forms.Panel();
            this.nudFoodPrice = new System.Windows.Forms.NumericUpDown();
            this.lbPrice = new System.Windows.Forms.Label();
            this.panel11 = new System.Windows.Forms.Panel();
            this.cbFoodCategory = new System.Windows.Forms.ComboBox();
            this.lbCategory = new System.Windows.Forms.Label();
            this.panel10 = new System.Windows.Forms.Panel();
            this.tbFoodsName = new System.Windows.Forms.TextBox();
            this.lbFoodName = new System.Windows.Forms.Label();
            this.panel9 = new System.Windows.Forms.Panel();
            this.tbFoodID = new System.Windows.Forms.TextBox();
            this.lbIDFood = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel8.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbfood)).BeginInit();
            this.panel15.SuspendLayout();
            this.panel12.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudFoodPrice)).BeginInit();
            this.panel11.SuspendLayout();
            this.panel10.SuspendLayout();
            this.panel9.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel8
            // 
            this.panel8.Controls.Add(this.label2);
            this.panel8.Controls.Add(this.bCancel);
            this.panel8.Controls.Add(this.bOK);
            this.panel8.Controls.Add(this.pbfood);
            this.panel8.Controls.Add(this.panel15);
            this.panel8.Controls.Add(this.panel12);
            this.panel8.Controls.Add(this.panel11);
            this.panel8.Controls.Add(this.panel10);
            this.panel8.Controls.Add(this.panel9);
            this.panel8.Location = new System.Drawing.Point(12, 12);
            this.panel8.Name = "panel8";
            this.panel8.Size = new System.Drawing.Size(525, 639);
            this.panel8.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label2.Location = new System.Drawing.Point(188, 221);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(172, 29);
            this.label2.TabIndex = 10;
            this.label2.Text = "Thêm hình ảnh";
            // 
            // bCancel
            // 
            this.bCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.bCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.bCancel.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.bCancel.Location = new System.Drawing.Point(318, 582);
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
            this.bOK.Location = new System.Drawing.Point(15, 582);
            this.bOK.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.bOK.Name = "bOK";
            this.bOK.Size = new System.Drawing.Size(174, 44);
            this.bOK.TabIndex = 6;
            this.bOK.Text = "OK";
            this.bOK.UseVisualStyleBackColor = false;
            this.bOK.Click += new System.EventHandler(this.bOK_Click);
            // 
            // pbfood
            // 
            this.pbfood.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pbfood.Location = new System.Drawing.Point(15, 13);
            this.pbfood.Name = "pbfood";
            this.pbfood.Size = new System.Drawing.Size(483, 220);
            this.pbfood.TabIndex = 9;
            this.pbfood.TabStop = false;
            this.pbfood.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // panel15
            // 
            this.panel15.Controls.Add(this.bFalse);
            this.panel15.Controls.Add(this.btrue);
            this.panel15.Controls.Add(this.label1);
            this.panel15.Location = new System.Drawing.Point(15, 440);
            this.panel15.Name = "panel15";
            this.panel15.Size = new System.Drawing.Size(483, 56);
            this.panel15.TabIndex = 5;
            // 
            // bFalse
            // 
            this.bFalse.AutoSize = true;
            this.bFalse.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.bFalse.Location = new System.Drawing.Point(372, 14);
            this.bFalse.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.bFalse.Name = "bFalse";
            this.bFalse.Size = new System.Drawing.Size(62, 24);
            this.bFalse.TabIndex = 2;
            this.bFalse.TabStop = true;
            this.bFalse.Text = "Hết ";
            this.bFalse.UseVisualStyleBackColor = true;
            // 
            // btrue
            // 
            this.btrue.AutoSize = true;
            this.btrue.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.btrue.Location = new System.Drawing.Point(246, 14);
            this.btrue.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btrue.Name = "btrue";
            this.btrue.Size = new System.Drawing.Size(60, 24);
            this.btrue.TabIndex = 1;
            this.btrue.TabStop = true;
            this.btrue.Text = "Còn";
            this.btrue.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label1.Location = new System.Drawing.Point(3, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(138, 29);
            this.label1.TabIndex = 0;
            this.label1.Text = "Trạng thái:";
            // 
            // panel12
            // 
            this.panel12.Controls.Add(this.nudFoodPrice);
            this.panel12.Controls.Add(this.lbPrice);
            this.panel12.Location = new System.Drawing.Point(15, 509);
            this.panel12.Name = "panel12";
            this.panel12.Size = new System.Drawing.Size(483, 56);
            this.panel12.TabIndex = 4;
            // 
            // nudFoodPrice
            // 
            this.nudFoodPrice.Location = new System.Drawing.Point(209, 14);
            this.nudFoodPrice.Maximum = new decimal(new int[] {
            -727379968,
            232,
            0,
            0});
            this.nudFoodPrice.Name = "nudFoodPrice";
            this.nudFoodPrice.Size = new System.Drawing.Size(254, 22);
            this.nudFoodPrice.TabIndex = 1;
            // 
            // lbPrice
            // 
            this.lbPrice.AutoSize = true;
            this.lbPrice.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lbPrice.Location = new System.Drawing.Point(3, 14);
            this.lbPrice.Name = "lbPrice";
            this.lbPrice.Size = new System.Drawing.Size(60, 29);
            this.lbPrice.TabIndex = 0;
            this.lbPrice.Text = "Giá:";
            // 
            // panel11
            // 
            this.panel11.Controls.Add(this.cbFoodCategory);
            this.panel11.Controls.Add(this.lbCategory);
            this.panel11.Location = new System.Drawing.Point(15, 377);
            this.panel11.Name = "panel11";
            this.panel11.Size = new System.Drawing.Size(483, 56);
            this.panel11.TabIndex = 3;
            // 
            // cbFoodCategory
            // 
            this.cbFoodCategory.FormattingEnabled = true;
            this.cbFoodCategory.Location = new System.Drawing.Point(209, 14);
            this.cbFoodCategory.Name = "cbFoodCategory";
            this.cbFoodCategory.Size = new System.Drawing.Size(254, 24);
            this.cbFoodCategory.TabIndex = 1;
            // 
            // lbCategory
            // 
            this.lbCategory.AutoSize = true;
            this.lbCategory.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lbCategory.Location = new System.Drawing.Point(3, 14);
            this.lbCategory.Name = "lbCategory";
            this.lbCategory.Size = new System.Drawing.Size(135, 29);
            this.lbCategory.TabIndex = 0;
            this.lbCategory.Text = "Danh mục:";
            // 
            // panel10
            // 
            this.panel10.Controls.Add(this.tbFoodsName);
            this.panel10.Controls.Add(this.lbFoodName);
            this.panel10.Location = new System.Drawing.Point(15, 315);
            this.panel10.Name = "panel10";
            this.panel10.Size = new System.Drawing.Size(483, 56);
            this.panel10.TabIndex = 2;
            // 
            // tbFoodsName
            // 
            this.tbFoodsName.Location = new System.Drawing.Point(209, 14);
            this.tbFoodsName.Name = "tbFoodsName";
            this.tbFoodsName.Size = new System.Drawing.Size(254, 22);
            this.tbFoodsName.TabIndex = 1;
            // 
            // lbFoodName
            // 
            this.lbFoodName.AutoSize = true;
            this.lbFoodName.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lbFoodName.Location = new System.Drawing.Point(3, 14);
            this.lbFoodName.Name = "lbFoodName";
            this.lbFoodName.Size = new System.Drawing.Size(158, 29);
            this.lbFoodName.TabIndex = 0;
            this.lbFoodName.Text = "Tên món ăn:";
            // 
            // panel9
            // 
            this.panel9.Controls.Add(this.tbFoodID);
            this.panel9.Controls.Add(this.lbIDFood);
            this.panel9.Location = new System.Drawing.Point(15, 253);
            this.panel9.Name = "panel9";
            this.panel9.Size = new System.Drawing.Size(483, 56);
            this.panel9.TabIndex = 1;
            // 
            // tbFoodID
            // 
            this.tbFoodID.Location = new System.Drawing.Point(209, 14);
            this.tbFoodID.Name = "tbFoodID";
            this.tbFoodID.Size = new System.Drawing.Size(254, 22);
            this.tbFoodID.TabIndex = 1;
            // 
            // lbIDFood
            // 
            this.lbIDFood.AutoSize = true;
            this.lbIDFood.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.lbIDFood.Location = new System.Drawing.Point(3, 14);
            this.lbIDFood.Name = "lbIDFood";
            this.lbIDFood.Size = new System.Drawing.Size(147, 29);
            this.lbIDFood.TabIndex = 0;
            this.lbIDFood.Text = "Mã món ăn:";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // AddFood
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LemonChiffon;
            this.ClientSize = new System.Drawing.Size(552, 687);
            this.Controls.Add(this.panel8);
            this.Name = "AddFood";
            this.Text = "AddFood";
            this.panel8.ResumeLayout(false);
            this.panel8.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbfood)).EndInit();
            this.panel15.ResumeLayout(false);
            this.panel15.PerformLayout();
            this.panel12.ResumeLayout(false);
            this.panel12.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudFoodPrice)).EndInit();
            this.panel11.ResumeLayout(false);
            this.panel11.PerformLayout();
            this.panel10.ResumeLayout(false);
            this.panel10.PerformLayout();
            this.panel9.ResumeLayout(false);
            this.panel9.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel8;
        private System.Windows.Forms.Button bCancel;
        private System.Windows.Forms.Button bOK;
        private System.Windows.Forms.Panel panel15;
        private System.Windows.Forms.RadioButton bFalse;
        private System.Windows.Forms.RadioButton btrue;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel12;
        private System.Windows.Forms.NumericUpDown nudFoodPrice;
        private System.Windows.Forms.Label lbPrice;
        private System.Windows.Forms.Panel panel11;
        private System.Windows.Forms.ComboBox cbFoodCategory;
        private System.Windows.Forms.Label lbCategory;
        private System.Windows.Forms.Panel panel10;
        private System.Windows.Forms.TextBox tbFoodsName;
        private System.Windows.Forms.Label lbFoodName;
        private System.Windows.Forms.Panel panel9;
        private System.Windows.Forms.TextBox tbFoodID;
        private System.Windows.Forms.Label lbIDFood;
        private System.Windows.Forms.PictureBox pbfood;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label2;
    }
}