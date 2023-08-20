using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.BLL;
using WindowsFormsApp1.DAO;
using WindowsFormsApp1.DTO;

namespace WindowsFormsApp1
{
    public partial class Menu : Form
    {
        private Table _table;

        public Table table
        {
            get { return _table; }
            set { _table = value; } 
        }
        private string _MaNhanVien;

        public string MaNhanVien
        {
            get { return _MaNhanVien; }
            set { _MaNhanVien = value; }
        }
        private string _MaMonAn;

        public string MaMonAn
        {
            get { return _MaMonAn; }
            set { _MaMonAn = value; }
        }

        private bool _TinhTrang;

        public Menu()
        {
            InitializeComponent();
            load();
        }

        #region methods
        public void GetData(Table table, string MaNhanVien)
        {
            this.table = table;
            this.MaNhanVien = MaNhanVien;
        }

        void load()
        {
            loadMenu();
        }

        void loadMenu()
        {
            flowLayoutPanel1.Controls.Clear();
            List<Category> Menulist = BLL_QLCF.Instance.GetAllCategories();
            foreach (Category item in Menulist)
            {
                Button bt = new Button() { Width = 150, Height = 150 };
                bt.Text = item.TenLoaiMonAn;
                bt.Font = new Font("Microsoft Sans Serif", 15f);
                bt.Click += bt_Click;
                bt.Tag = item; 
                bt.BackColor = Color.LightGreen;
                flowLayoutPanel1.Controls.Add(bt);
            }
            buttonAdd.Hide();
            nudThem.Hide();
        }

        void loadFood(string MaLoaiMonAn)
        {
            flowLayoutPanel1.Controls.Clear();
            List<DoAn> Menulist = BLL_QLCF.Instance.GetFoodByCategory(MaLoaiMonAn);
            foreach (DoAn item in Menulist)
            {
                Button bt = new Button() { Width = 150, Height = 150 };
                if(item.HinhAnh != null)
                {
                    MemoryStream ms = new MemoryStream(item.HinhAnh);
                    Image image = Image.FromStream(ms);
                    bt.BackgroundImage = image;
                    bt.BackgroundImageLayout = ImageLayout.Stretch;
                }
                bt.Click += bt2_Click;
                bt.Tag = item; 
                Label label = new Label();
                label.Text = item.TenMonAn;
                label.Font = new Font("Microsoft Sans Serif", 11f);
                CustomControl1 c = new CustomControl1(bt, label);
                flowLayoutPanel1.Controls.Add(c);
            }
        }
        #endregion

        #region events
        private void bt2_Click(object sender, EventArgs e)
        {
            this.MaMonAn = ((sender as Button).Tag as DoAn).MaMonAn;
            this._TinhTrang = ((sender as Button).Tag as DoAn).TinhTrangMonAn;
            buttonAdd.Show();
            nudThem.Show();
        }

        private void bt_Click(object sender, EventArgs e)
        {
            string MaLoaiMonAn = ((sender as Button).Tag as Category).MaLoaiMonAn;
            loadFood(MaLoaiMonAn);
        }

        private void btRefresh_Click(object sender, EventArgs e)
        {
            loadMenu();
            buttonAdd.Hide();
            nudThem.Hide();
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            string MaHoaDon = BLL_QLCF.Instance.GetUncheckBillIDByTableID(table.MaBan);
            string MaMonAn = this.MaMonAn;
            int SoLuong = (int)nudThem.Value;

            if (_TinhTrang == false)
            {
                MessageBox.Show("Món này hiện tại đã hết!!!", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MaHoaDon.Equals("-1"))
            {
                if (SoLuong <= 0)
                {
                    MessageBox.Show("Số lượng không hợp lệ!!!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                MaHoaDon = BLL_QLCF.Instance.GetNewBillID();
                BLL_QLCF.Instance.InsertIntoBill(MaHoaDon, table.MaBan, MaNhanVien);
            }

            if (BLL_QLCF.Instance.InsertBillData(MaHoaDon, MaMonAn, SoLuong))
            {
                MessageBox.Show("Thêm món thành công", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Số lượng không hợp lệ!!!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {


        }
    }
}
