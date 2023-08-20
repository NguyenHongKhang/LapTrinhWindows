
namespace WindowsFormsApp1.Views
{
    partial class AddCategory
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
            this.panel48 = new System.Windows.Forms.Panel();
            this.tbCategoryName = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.bCancel = new System.Windows.Forms.Button();
            this.bOK = new System.Windows.Forms.Button();
            this.panel49 = new System.Windows.Forms.Panel();
            this.tbCategoryID = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.panel48.SuspendLayout();
            this.panel49.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel48
            // 
            this.panel48.Controls.Add(this.tbCategoryName);
            this.panel48.Controls.Add(this.label19);
            this.panel48.Location = new System.Drawing.Point(12, 78);
            this.panel48.Name = "panel48";
            this.panel48.Size = new System.Drawing.Size(483, 56);
            this.panel48.TabIndex = 3;
            // 
            // tbCategoryName
            // 
            this.tbCategoryName.Location = new System.Drawing.Point(226, 14);
            this.tbCategoryName.Name = "tbCategoryName";
            this.tbCategoryName.Size = new System.Drawing.Size(254, 22);
            this.tbCategoryName.TabIndex = 1;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label19.Location = new System.Drawing.Point(3, 14);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(208, 29);
            this.label19.TabIndex = 0;
            this.label19.Text = "Tên loại món ăn:";
            // 
            // bCancel
            // 
            this.bCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.bCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.bCancel.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.bCancel.Location = new System.Drawing.Point(315, 149);
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
            this.bOK.Location = new System.Drawing.Point(12, 149);
            this.bOK.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.bOK.Name = "bOK";
            this.bOK.Size = new System.Drawing.Size(174, 44);
            this.bOK.TabIndex = 9;
            this.bOK.Text = "OK";
            this.bOK.UseVisualStyleBackColor = false;
            this.bOK.Click += new System.EventHandler(this.bOK_Click);
            // 
            // panel49
            // 
            this.panel49.Controls.Add(this.tbCategoryID);
            this.panel49.Controls.Add(this.label20);
            this.panel49.Location = new System.Drawing.Point(12, 16);
            this.panel49.Name = "panel49";
            this.panel49.Size = new System.Drawing.Size(483, 56);
            this.panel49.TabIndex = 11;
            // 
            // tbCategoryID
            // 
            this.tbCategoryID.Location = new System.Drawing.Point(226, 12);
            this.tbCategoryID.Name = "tbCategoryID";
            this.tbCategoryID.ReadOnly = true;
            this.tbCategoryID.Size = new System.Drawing.Size(254, 22);
            this.tbCategoryID.TabIndex = 1;
            this.tbCategoryID.TextChanged += new System.EventHandler(this.tbCategoryID_TextChanged);
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.label20.Location = new System.Drawing.Point(3, 14);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(197, 29);
            this.label20.TabIndex = 0;
            this.label20.Text = "Mã loại món ăn:";
            // 
            // AddCategory
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LemonChiffon;
            this.ClientSize = new System.Drawing.Size(504, 207);
            this.Controls.Add(this.panel49);
            this.Controls.Add(this.bCancel);
            this.Controls.Add(this.bOK);
            this.Controls.Add(this.panel48);
            this.Name = "AddCategory";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Thêm loại món ăn";
            this.panel48.ResumeLayout(false);
            this.panel48.PerformLayout();
            this.panel49.ResumeLayout(false);
            this.panel49.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel48;
        private System.Windows.Forms.TextBox tbCategoryName;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Button bCancel;
        private System.Windows.Forms.Button bOK;
        private System.Windows.Forms.Panel panel49;
        private System.Windows.Forms.TextBox tbCategoryID;
        private System.Windows.Forms.Label label20;
    }
}