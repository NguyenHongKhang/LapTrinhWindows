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

namespace WindowsFormsApp1
{
    public partial class AccountProfile : Form
    {
        private Account _account;

        public Account account
        {
            get { return _account; }
            set { _account = value; SetAccount(account); }
        }

        public AccountProfile(Account account)
        {
            InitializeComponent();

            this.account = account;
        }
        void SetAccount(Account account)
        {
            tbUserName.Text = account.TenTaiKhoan;
            tbAppealName.Text = account.TenHienThi;
        }
        void updateAccount()
        {
            Account ac = new Account(tbUserName.Text, tbPassword.Text, tbAppealName.Text, account.LoaiTaiKhoan, account.MaNhanVien);
            string MatKhauMoi = tbMatKhauMoi.Text;
            string XacNhanMKM = tbNhapLaiMatKhau.Text;
            bool result;
            try
            {
                result = BLL_QLCF.Instance.UpdateAccount(ac, MatKhauMoi, XacNhanMKM);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (result == false)
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình cập nhật", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Cập nhật thành công!!!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return;
        }
        private void buttonExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonUpdate_Click(object sender, EventArgs e)
        {
            updateAccount();
            this.Close();
        }

        private void AccountProfile_Load(object sender, EventArgs e)
        {


        }
    }
}
