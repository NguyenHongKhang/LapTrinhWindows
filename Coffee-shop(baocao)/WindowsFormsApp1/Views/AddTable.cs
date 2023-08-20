using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.BLL;
using WindowsFormsApp1.DTO;

namespace WindowsFormsApp1.Views
{
    public partial class AddTable : Form
    {
        public AddTable()
        {
            InitializeComponent();
            tbMaBan.Text = BLL_QLCF.Instance.GetNewTableId();
            tbTrangThai.Text = "Trống";
        }
        private bool InsertTable()
        {
            Table kh = new Table(tbMaBan.Text, tbTenBan.Text, tbTrangThai.Text);
            bool result;
            try
            {
                result = BLL_QLCF.Instance.InsertTable(kh);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (result)
            {
                MessageBox.Show("Thêm bàn thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình thêm bàn", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return true;
        }
        private void bOK_Click(object sender, EventArgs e)
        {
            if (InsertTable() == true)
                this.Dispose();
            else return;
        }

        private void bCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }

}
