using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.BLL;
using WindowsFormsApp1.DAL;
using WindowsFormsApp1.DAO;
using WindowsFormsApp1.DTO;
using WindowsFormsApp1.Views;
using static System.Windows.Forms.ListViewItem;
using Bill = WindowsFormsApp1.DTO.Bill;

namespace WindowsFormsApp1
{
    public delegate void PassingData(Table table, string MaNhanVien);
    public partial class TableManagement : Form
    {
        private Account _account;

        public Account account 
        {
            get { return _account; }
            set { _account = value; changeAccount(account.LoaiTaiKhoan); }
        }

        public TableManagement(Account account)
        {
            InitializeComponent();
            this.account = account;
            loadTable();
            LoadTableToComboBox();
        }

        #region Method

        void changeAccount(int LoaiTaiKhoan)
        {
            adminToolStripMenuItem.Visible = LoaiTaiKhoan == 1;
        }

        void loadTable()
        {
            List<Table> tablelist = BLL_QLCF.Instance.GetAllTable();
            foreach (Table item in tablelist)
            {
                Button bt = new Button() {Width = 90, Height = 90 };
                bt.Text = item.TenBan + Environment.NewLine + item.TrangThai;
                bt.Click += bt_Click;
                bt.Tag = item; 
                switch (item.TrangThai)
                {
                    case "Trống":
                        bt.BackColor = Color.LightBlue;
                        break;
                    default:
                        bt.BackColor = Color.LightGreen;
                        break;
                }
                flpTable.Controls.Add(bt);
            }
        }

        void ShowBill(string MaBan)
        {
            lsvBill.Items.Clear();
            List<Bill> listChiTietHoaDon = BLL_QLCF.Instance.GetListBillByTable(MaBan);
            float tongtien = 0;
            foreach (Bill item in listChiTietHoaDon)
            {
                ListViewItem lsvitem = new ListViewItem(item.TenMonAn.ToString());
                lsvitem.SubItems.Add(item.SoLuong.ToString());
                lsvitem.SubItems.Add(item.Price.ToString());
                lsvitem.SubItems.Add(item.ThanhTien.ToString());
                tongtien += item.ThanhTien;
                lsvBill.Items.Add(lsvitem);
            }
            CultureInfo culture = new CultureInfo("vn-VN");
            tbTongTien.Text = tongtien.ToString("c", culture);
        }

        void LoadTableToComboBox()
        {
            cbMoveTable.DataSource = BLL_QLCF.Instance.GetAllTable();
            cbMoveTable.DisplayMember = "TenBan";
        }

        #endregion

        #region Events

        private void bt_Click(object sender, EventArgs e)
        {
            string MaBan = ((sender as Button).Tag as Table).MaBan;
            lsvBill.Tag = (sender as Button).Tag;
            ShowBill(MaBan);
        }
        private void buttonAdd_Click(object sender, EventArgs e)
        {
            Table table = lsvBill.Tag as Table;
            if (table == null)
            {
                MessageBox.Show("Hãy chọn bàn trước", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Menu f2 = new Menu();
            PassingData pd = new PassingData(f2.GetData);
            pd(table,account.MaNhanVien);
            f2.ShowDialog();
            flpTable.Controls.Clear();
            ShowBill(table.MaBan);
            loadTable();
            LoadTableToComboBox();
        }

        private void buttonPay_Click(object sender, EventArgs e)
        {
            Table table = lsvBill.Tag as Table;
            if (table == null)
            {
                MessageBox.Show("Hãy chọn bàn trước", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string MaHoaDon = BLL_QLCF.Instance.GetUncheckBillIDByTableID(table.MaBan);
            if (!MaHoaDon.Equals("-1"))
            {
                DialogResult dialogResult = MessageBox.Show("Bạn có muốn sử dụng tài khoản không ?", "Thông báo", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    string formattedTongTien = tbTongTien.Text.Replace("₫", ""); // Xóa ký tự đơn vị tiền tệ
                    float tongTien = float.Parse(formattedTongTien.Replace(",", ""), CultureInfo.InvariantCulture); // Chuyển đổi chuỗi thành số thực
                    int intTongTien = (int)tongTien; // Chuyển đổi số thực thành số nguyên
                    Payments f2 = new Payments(MaHoaDon, intTongTien, table.MaBan, account.MaNhanVien);
                    f2.ShowDialog();
                }
                else if (dialogResult == DialogResult.No)
                {
                    string formattedTongTien = tbTongTien.Text.Replace("₫", ""); // Xóa ký tự đơn vị tiền tệ
                    float tongTien = float.Parse(formattedTongTien.Replace(",", ""), CultureInfo.InvariantCulture); // Chuyển đổi chuỗi thành số thực
                    int intTongTien = (int)tongTien; // Chuyển đổi số thực thành số nguyên
                    Payments2 f2 = new Payments2(MaHoaDon, intTongTien, table.MaBan, account.MaNhanVien);
                    f2.ShowDialog();
                }

                flpTable.Controls.Clear();
                ShowBill(table.MaBan);
                loadTable();
                LoadTableToComboBox();
            }
        }




        private void btCancelTable_Click(object sender, EventArgs e)
        {
            Table table = lsvBill.Tag as Table;
            if (table == null)
            {
                MessageBox.Show("Hãy chọn bàn trước", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string MaHoaDon = BLL_QLCF.Instance.GetUncheckBillIDByTableID(table.MaBan);
            if (MaHoaDon.Equals("-1") == false)
            {
                if (MessageBox.Show("Bạn có thực sự muốn hủy không ?", "Thông báo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    DAL_QLCF.Instance.TableCancel(MaHoaDon, table.MaBan);
                }
                flpTable.Controls.Clear();
                ShowBill(table.MaBan);
                loadTable();
                LoadTableToComboBox();
            }
        }

        private void btMoveTable_Click(object sender, EventArgs e)
        {
            Table table1 = lsvBill.Tag as Table;
            Table table2 = (Table)cbMoveTable.SelectedItem;

            if (table1 == null || table2 == null)
            {
                MessageBox.Show("Mời chọn bàn để chuyển!!!", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                if (BLL_QLCF.Instance.MoveTable(table1, table2) == false)
                {
                    MessageBox.Show("Có lỗi xảy ra!!!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("Chuyển bàn thành công!!!", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    flpTable.Controls.Clear();
                    ShowBill(table1.MaBan);
                    loadTable();
                    LoadTableToComboBox();
                    lsvBill.Tag = null;
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btTableMerge_Click(object sender, EventArgs e)
        {
            Table table1 = lsvBill.Tag as Table;
            Table table2 = (Table)cbMoveTable.SelectedItem;

            if (table1 == null || table2 == null)
            {
                MessageBox.Show("Mời chọn bàn để gộp!!!", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                if (BLL_QLCF.Instance.MergeTable(table1, table2) == false)
                {
                    MessageBox.Show("Có lỗi xảy ra!!!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("Gộp bàn thành công!!!", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    flpTable.Controls.Clear();
                    ShowBill(table1.MaBan);
                    loadTable();
                    LoadTableToComboBox();
                    lsvBill.Tag = null;
                }
            }

            catch(Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void tsmiQuanly_Click(object sender, EventArgs e)
        {
            Admin f = new Admin();
            f.ShowDialog();
            flpTable.Controls.Clear();
            loadTable();
        }

        private void tsmiThongke_Click(object sender, EventArgs e)
        {
            Statistics f = new Statistics();
            f.ShowDialog();
            flpTable.Controls.Clear();
            loadTable();
        }

        private void tsmiThongtin_Click(object sender, EventArgs e)
        {
            NhanVien nhanvien = BLL_QLCF.Instance.GetStaffByStaffId(account.MaNhanVien);
            AccountView p = new AccountView(nhanvien);
            p.Show();
        }

        private void tsmiThaydoi_Click(object sender, EventArgs e)
        {
            AccountProfile f = new AccountProfile(account);
            f.ShowDialog();
        }

        private void tsmiDangxuat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void tsmiKhachhang_Click(object sender, EventArgs e)
        {
            Customer f = new Customer();
            f.ShowDialog();
        }

        #endregion

        private void tbTongTien_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lsvBill_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void TableManagement_Load(object sender, EventArgs e)
        {

        }

        private void flpTable_Paint(object sender, PaintEventArgs e)
        {


        }
    }
}
