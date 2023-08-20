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
using WindowsFormsApp1.DAL;
using WindowsFormsApp1.DAO;
using WindowsFormsApp1.DTO;
using WindowsFormsApp1.Views;

namespace WindowsFormsApp1
{
    public partial class Customer : Form
    {
        BindingSource dgvCustomerBinding = new BindingSource();
        public Customer()
        {
            InitializeComponent();
            load();
            dgvCustomerBinding.DataSource = BLL_QLCF.Instance.GetAllCustomer();
        }
        void load()
        {
            LoadListCustomer();
            addCustomerBinding();
        }
        void LoadListCustomer()
        {
            dgvCustomer.DataSource = dgvCustomerBinding;
        }
        private void addCustomerBinding()
        {
            tbMaKhachHang.DataBindings.Add(new Binding("Text", dgvCustomer.DataSource, "MaKhachHang", true, DataSourceUpdateMode.Never));
            tbTenKhachHang.DataBindings.Add(new Binding("Text", dgvCustomer.DataSource, "TenKhachHang", true, DataSourceUpdateMode.Never));
            tbSoDienThoai.DataBindings.Add(new Binding("Text", dgvCustomer.DataSource, "SoDienThoai", true, DataSourceUpdateMode.Never));
            tbSoBillsDaThanhToan.DataBindings.Add(new Binding("Text", dgvCustomer.DataSource, "SoBillsDaThanhToan", true, DataSourceUpdateMode.Never));
            tbDiemTichLuy.DataBindings.Add(new Binding("Text", dgvCustomer.DataSource, "DiemTichLuy", true, DataSourceUpdateMode.Never));
        }
        private void ShowCustomer()
        {
            dgvCustomerBinding.DataSource = BLL_QLCF.Instance.FindCustomer(tbfind.Text);
        }
        private void buttonAddCustomer_Click(object sender, EventArgs e)
        {
            AddCustomer f2 = new AddCustomer();
            f2.ShowDialog();
            ShowCustomer();
        }

        private void btfind_Click(object sender, EventArgs e)
        {
            ShowCustomer();
        }

        private void buttonDeleteCustomer_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có thật sự muốn xóa " +dgvCustomer.SelectedRows.Count.ToString() + " khách hàng?", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                DeleteCustomer();
                ShowCustomer();
            }
        }
        private void DeleteCustomer()
        {
            List<string> ListDelete = new List<string>();
            foreach (DataGridViewRow i in dgvCustomer.SelectedRows)
            {
                ListDelete.Add(i.Cells["MaKhachHang"].Value.ToString());
            }
            if (BLL_QLCF.Instance.DeleteCustomer(ListDelete))
            {
                MessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình xóa", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonFixCustomer_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có thật sự muốn cập nhập thông tin khách hàng?", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                UpdateCustomer();
                ShowCustomer();
            }
        }
        private void UpdateCustomer()
        {
            string tenKhachHang = tbTenKhachHang.Text;
            string maKhachHang = tbMaKhachHang.Text;
            string soDienThoai = tbSoDienThoai.Text;
            int soBillsDaThanhToan = Convert.ToInt32(tbSoBillsDaThanhToan.Text);

            int diemTichLuy;
            if (!int.TryParse(tbDiemTichLuy.Text, out diemTichLuy))
            {
                MessageBox.Show("Vui lòng nhập số vào trường Điểm tích luỹ.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            KhachHang kh = new KhachHang(tenKhachHang, maKhachHang, soDienThoai, soBillsDaThanhToan, diemTichLuy);

            bool result;
            try
            {
                result = BLL_QLCF.Instance.UpdateCustomer(kh);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (result)
            {
                MessageBox.Show("Cập nhật thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình cập nhật", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {


        }

        private void tbMaKhachHang_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
