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
using WindowsFormsApp1.DAO;
using WindowsFormsApp1.DTO;

namespace WindowsFormsApp1
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();           
        }

        private void buttonLogin_Click(object sender, EventArgs e)
        {
            string TenTaiKhoan = tbUserName.Text;
            string MatKhau = tbPassword.Text;
            Account ac = BLL_QLCF.Instance.GetAccount(TenTaiKhoan, MatKhau);
            if (ac != null)
            {
                TableManagement f = new TableManagement(ac);
                this.Hide();
                f.ShowDialog();
                this.Show();
            }
            else
            {
                MessageBox.Show("Sai tên tài khoản hoặc mật khẩu !!!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbPassword.Text = "";
            }
            tbPassword.Text = "";
        }
        private void buttonExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Login_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Bạn có thật sự muốn thoát ?","Cảnh báo", MessageBoxButtons.OKCancel,MessageBoxIcon.Question) != DialogResult.OK)
            {
                e.Cancel = true;
            }
        }

        private void Login_Load(object sender, EventArgs e)
        {


        }
    }
}
