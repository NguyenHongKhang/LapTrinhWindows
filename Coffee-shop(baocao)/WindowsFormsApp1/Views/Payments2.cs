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
    public partial class Payments2 : Form
    {
        private string _MaHoaDon;

        public string MaHoaDon
        {
            get { return _MaHoaDon; }
            set { _MaHoaDon = value; }
        }
        private float _TamTinh;

        public float TamTinh
        {
            get { return _TamTinh; }
            set { _TamTinh = value; }
        }
        private string _MaBan;

        public string MaBan
        {
            get { return _MaBan; }
            set { _MaBan = value; }
        }
        private string _MaNhanVien;

        public string MaNhanVien
        {
            get { return _MaNhanVien; }
            set { _MaNhanVien = value; }
        }
        public Payments2(string MaHoaDon, float TamTinh, string MaBan, string MaNhanVien)
        {
            InitializeComponent();
            this.MaHoaDon = MaHoaDon;
            this.TamTinh = TamTinh;
            this.MaBan = MaBan;
            this.MaNhanVien = MaNhanVien;
            load();
        }
        #region method
        void load()
        {
            setCBBGiamGia();
            setTB();
        }
        void setCBBGiamGia()
        {
            cbGiamGia.Items.Add(new CBBItems { Text = "0%", Value = 0 });
            cbGiamGia.Items.Add(new CBBItems { Text = "10%", Value = 1 });
            cbGiamGia.Items.Add(new CBBItems { Text = "15%", Value = 2 });
            cbGiamGia.SelectedIndex = 0;
        }
        void setTB()
        {
            tbTongTien.Text = TamTinh.ToString();
        }
        #endregion

        private void cbGiamGia_SelectedIndexChanged(object sender, EventArgs e)
        {
            float TongTien = TamTinh;
            if (cbGiamGia.SelectedIndex == 1)
            {
                TongTien = TamTinh - TamTinh / 10;
            }
            if (cbGiamGia.SelectedIndex == 2)
            {
                TongTien = TamTinh - TamTinh * 15 / 100;
            }
            tbSauGiamGia.Text = TongTien.ToString();
            if (tbSoTienKhachTra.Text != "")
            {
                float TienThoi = float.Parse(tbSoTienKhachTra.Text) - float.Parse(tbSauGiamGia.Text);
                tbSoTienThoiLai.Text = TienThoi.ToString();
            }
        }

        

        private void btThanhToan_Click(object sender, EventArgs e)
        {
            AddressShop shop = BLL_QLCF.Instance.GetShop();
            NhanVien nhanVien = BLL_QLCF.Instance.GetStaffByStaffId(MaNhanVien);
            try
            {
                float TienThoi = float.Parse(tbSoTienThoiLai.Text);
                if (TienThoi < 0)
                {
                    throw new InvalidOperationException("Số tiền nhập vào không hợp lệ");
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (MessageBox.Show("Bạn có chắc muốn thanh toán không ?", "Thông báo", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                string newPath = @"E:\BT C#\Coffee-shop(baocao)\Hoá đơn";
                string wr = "";
                System.IO.FileStream fs = new System.IO.FileStream(newPath + "\\" + MaHoaDon + ".txt", FileMode.Create, FileAccess.Write, FileShare.None);

                StreamWriter sw = new StreamWriter(fs);
                sw.WriteLine("                     " + shop.TenShop + "                     \n");
                sw.WriteLine("Mã nhân viên     : " + MaNhanVien + "\n");
                sw.WriteLine("Tên nhân viên    : " + nhanVien.TenNhanVien + "\n");
                sw.WriteLine("Mã hóa đơn   : " + MaHoaDon + "\n");
                sw.WriteLine("--------------------------------------------------\n");
                wr += string.Format("{0,-20}", "Tên món");
                wr += string.Format("{0,10}", "Số lượng");
                wr += string.Format("{0,17}", "Thành tiền");
                sw.WriteLine(wr + "\n");
                List<Bill> listChiTietHoaDon = BLL_QLCF.Instance.GetListBillByTable(MaBan);
                foreach (Bill item in listChiTietHoaDon)
                {
                    string wr2 = "";
                    wr2 += string.Format("{0,-20}", item.TenMonAn.ToString());
                    wr2 += string.Format("{0,0}", "|");
                    wr2 += string.Format("{0,5}", item.SoLuong.ToString());
                    wr2 += string.Format("{0,7}", "|");
                    wr2 += string.Format("{0,10:C}", item.ThanhTien.ToString());
                    sw.WriteLine(wr2 + "\n");
                }
                sw.WriteLine("--------------------------------------------------\n");
                string wr4 = "";
                wr4 += string.Format("{0,-20}", "Tạm tính");
                wr4 += string.Format("{0,20}", (tbTongTien.Text.ToString()));
                string wr3 = "";
                wr3 += string.Format("{0,-20}", "Giảm giá");
                wr3 += string.Format("{0,20}", ((CBBItems)cbGiamGia.SelectedItem).Text.ToString());
                string wr5 = "";
                wr5 += string.Format("{0,-20}", "Tổng tiền");
                wr5 += string.Format("{0,20}", tbSauGiamGia.Text.ToString());
                string wr6 = "";
                wr6 += string.Format("{0,-20}", "Số tiền khách trả");
                wr6 += string.Format("{0,20}", tbSoTienKhachTra.Text.ToString());
                string wr7 = "";
                wr7 += string.Format("{0,-20}", "Số tiền thối lại");
                wr7 += string.Format("{0,20}", tbSoTienThoiLai.Text.ToString());
                sw.WriteLine(wr4 + "\n");
                sw.WriteLine(wr3 + "\n");
                sw.WriteLine(wr5 + "\n");
                sw.WriteLine(wr6 + "\n");
                sw.WriteLine(wr7 + "\n");
                sw.WriteLine("--------------------------------------------------\n");
                string wr8 = "";
                wr8 += string.Format("{0,-20}", "Thời gian");
                wr8 += string.Format("{0,17}", DateTime.Now.ToString());
                sw.WriteLine(wr8 + "\n");
                string wr11 = "";
                wr11 += string.Format("{0,-20}", "Tên chủ shop");
                wr11 += string.Format("{0,17}", shop.TenChuShop);
                sw.WriteLine(wr11 + "\n");
                string wr9 = "";
                wr9 += string.Format("{0,-20}", "Địa chỉ");
                wr9 += string.Format("{0,17}", shop.DiaChiShop);
                sw.WriteLine(wr9 + "\n");
                string wr10 = "";
                wr10 += string.Format("{0,-20}", "Số điện thoại liên hệ");
                wr10 += string.Format("{0,17}", shop.SoDienThoai);
                sw.WriteLine(wr10 + "\n");
                sw.Flush();
                sw.Close();
                fs.Close();
                BLL_QLCF.Instance.ChechOut(MaHoaDon, "", ((CBBItems)cbGiamGia.SelectedItem).Text, tbSauGiamGia.Text, tbSoTienKhachTra.Text, tbSoTienThoiLai.Text); ; ;
                this.Dispose();
            }
        }

        private void tbSoTienKhachTra_Leave(object sender, EventArgs e)
        {
            try
            {
                float TienThoi = float.Parse(tbSoTienKhachTra.Text) - float.Parse(tbSauGiamGia.Text);
                tbSoTienThoiLai.Text = TienThoi.ToString();
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {


        }
    }
}
