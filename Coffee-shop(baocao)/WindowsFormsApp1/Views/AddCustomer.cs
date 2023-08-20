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
    public partial class AddCustomer : Form
    {
        public AddCustomer()
        {
            InitializeComponent();
            IDKhachHang.Text = BLL_QLCF.Instance.GetNewCustomerId();
        }
        private bool InsertCustomer()
        {
            //Tạo biến id tăng tự động
            KhachHang kh = new KhachHang(tbTenKhachHang.Text, IDKhachHang.Text, tbSDT.Text, 0, 0);
            bool result;
            try
            {
                result = BLL_QLCF.Instance.InsertCustomer(kh);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (result)
            {
                MessageBox.Show("Thêm Khách hàng thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình Khách hàng", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return true; 
        }
        private void bCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void bOK_Click(object sender, EventArgs e)
        {
            if (InsertCustomer() == true)
                this.Dispose();
            else return;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
