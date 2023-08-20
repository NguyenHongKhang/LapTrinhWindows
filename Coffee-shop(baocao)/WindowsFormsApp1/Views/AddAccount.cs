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
    public partial class AddAccount : Form
    {
        public AddAccount()
        {
            InitializeComponent();
            LoadForm();
        }
        private void GetStaffIDToComboBox()
        {
            cbMaNhanVien.DataSource = BLL_QLCF.Instance.GetAllStaff();
            cbMaNhanVien.DisplayMember = "MaNhanVien";
            cbLoaiTaiKhoan.DataSource = BLL_QLCF.Instance.GetAccountTypeCBBItem();
        }
        private void LoadForm()
        {
            GetStaffIDToComboBox();
        }
        private bool InsertAccount()
        {
            Account da = new Account(tbTenTaiKhoan.Text,tbMatKhau.Text,tbTenHienThi.Text,((CBBItems)cbLoaiTaiKhoan.SelectedItem).Value,cbMaNhanVien.Text);
            bool result;
            try
            {
                result = BLL_QLCF.Instance.InsertAccount(da);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (result)
            {
                MessageBox.Show("Thêm tài khoản thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình thêm tài khoản", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return true;
        }
        private void bOK_Click(object sender, EventArgs e)
        {
            if (InsertAccount() == true)
                this.Dispose();
            else return;
        }

        private void bCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void cbLoaiTaiKhoan_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tbTenTaiKhoan_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
