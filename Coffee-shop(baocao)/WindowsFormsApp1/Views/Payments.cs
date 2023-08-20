using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.BLL;
using WindowsFormsApp1.DAL;
using WindowsFormsApp1.DAO;
using WindowsFormsApp1.DTO;
using System.Globalization;
using iTextSharp.text;
using iTextSharp.text.pdf;
using OfficeOpenXml;
using System.Xml.Linq;

namespace WindowsFormsApp1
{
    public partial class Payments : Form
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
        BindingSource DgvCustomerBinding;

        public Payments(string MaHoaDon,float TamTinh, string MaBan, string MaNhanVien)
        {
            InitializeComponent();
            DgvCustomerBinding = new BindingSource();
            dgvCustomer.DataSource = DgvCustomerBinding;
            this.MaHoaDon = MaHoaDon;
            this.TamTinh = TamTinh;
            this.MaBan = MaBan;
            this.MaNhanVien = MaNhanVien;
            load();
        }

        

        #region method
        void load()
        {
            LoadCustomer();
            setCBBGiamGia();
            setTB();
            addCustomerBinding();
        }

        void LoadCustomer()
        {
            DgvCustomerBinding.DataSource = BLL_QLCF.Instance.GetAllCustomer();
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

        #region events
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
            if ((int)Convert.ToInt32(tbDiemTichLuy.Text) < 10 && cbGiamGia.SelectedIndex > 0)
            {
                MessageBox.Show("Khách hàng không đủ điểm để giảm giá", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if ((int)Convert.ToInt32(tbDiemTichLuy.Text) < 20 && cbGiamGia.SelectedIndex > 1)
            {
                MessageBox.Show("Khách hàng không đủ điểm để giảm giá", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if(tbSoTienKhachTra.Text == "")
            {
                MessageBox.Show("Vui lòng nhập vào số tiền khách trả", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show("Bạn có chắc muốn thanh toán không ?", "Thông báo", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                string newPath = @"E:\BT C#\Coffee-shop(baocao)\Hoá đơn\TXT";
                string wr = "";
                System.IO.FileStream fs = new System.IO.FileStream(newPath + "\\" + MaHoaDon + ".txt", FileMode.Create, FileAccess.Write, FileShare.None);

                StreamWriter sw = new StreamWriter(fs);
                sw.WriteLine("                     "+shop.TenShop+"                     \n");
                sw.WriteLine("Mã nhân viên     : " + MaNhanVien + "\n");
                sw.WriteLine("Tên nhân viên    : " + nhanVien.TenNhanVien + "\n");
                sw.WriteLine("Mã hóa đơn       : " + MaHoaDon + "\n");
                sw.WriteLine("Mã khách hàng    : " + tbMaKhachHang.Text + "\n");
                sw.WriteLine("Tên khách hàng   : " + tbTenKhachHang.Text + "\n");
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
                MessageBox.Show("Thanh toán thành công.");
                sw.Flush();
                sw.Close();
                fs.Close();
                BLL_QLCF.Instance.ChechOut(MaHoaDon,tbMaKhachHang.Text,((CBBItems)cbGiamGia.SelectedItem).Text,tbSauGiamGia.Text,tbSoTienKhachTra.Text,tbSoTienThoiLai.Text);
                this.Dispose();
            }
        }

        void addCustomerBinding()
        {
            tbMaKhachHang.DataBindings.Add(new Binding("Text", dgvCustomer.DataSource, "MaKhachHang"));
            tbTenKhachHang.DataBindings.Add(new Binding("Text", dgvCustomer.DataSource, "TenKhachHang"));
            tbSoDienThoai.DataBindings.Add(new Binding("Text", dgvCustomer.DataSource, "SoDienThoai"));
            tbSoBillsDaThanhToan.DataBindings.Add(new Binding("Text", dgvCustomer.DataSource, "SoBillsDaThanhToan"));
            tbDiemTichLuy.DataBindings.Add(new Binding("Text", dgvCustomer.DataSource, "DiemTichLuy"));
        }

        private void cbGiamGia_SelectedIndexChanged(object sender, EventArgs e)
        {
            float TongTien = TamTinh;
            if (cbGiamGia.SelectedIndex == 1)
            {
                TongTien = TamTinh-TamTinh/10;
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

        private void btSearch_Click(object sender, EventArgs e)
        {
            dgvCustomer.DataSource = BLL_QLCF.Instance.GetListCustomer(tbCustomerName.Text);
        }

        #endregion

        private void btnExcel_Click(object sender, EventArgs e)
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

            if ((int)Convert.ToInt32(tbDiemTichLuy.Text) < 10 && cbGiamGia.SelectedIndex > 0)
            {
                MessageBox.Show("Khách hàng không đủ điểm để giảm giá", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if ((int)Convert.ToInt32(tbDiemTichLuy.Text) < 20 && cbGiamGia.SelectedIndex > 1)
            {
                MessageBox.Show("Khách hàng không đủ điểm để giảm giá", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (tbSoTienKhachTra.Text == "")
            {
                MessageBox.Show("Vui lòng nhập vào số tiền khách trả", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show("Bạn có chắc muốn thanh toán không ?", "Thông báo", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                string newPath = @"E:\BT C#\Coffee-shop(baocao)\Hoá đơn\EXCEL";
                Directory.CreateDirectory(newPath);
                string filePath = Path.Combine(newPath, MaHoaDon + ".xlsx");

                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("HoaDon");

                    // Populate Excel with data
                    int row = 1;
                    worksheet.Cells[row, 1].Value = "Tên món";
                    worksheet.Cells[row, 2].Value = "Số lượng";
                    worksheet.Cells[row, 3].Value = "Thành tiền";

                    row++;
                    List<Bill> listChiTietHoaDon = BLL_QLCF.Instance.GetListBillByTable(MaBan);
                    foreach (Bill item in listChiTietHoaDon)
                    {
                        worksheet.Cells[row, 1].Value = item.TenMonAn;
                        worksheet.Cells[row, 2].Value = item.SoLuong;
                        worksheet.Cells[row, 3].Value = item.ThanhTien;
                        row++;
                    }

                    // Additional information
                    worksheet.Cells[row + 1, 1].Value = "Tên shop:";
                    worksheet.Cells[row + 1, 2].Value = shop.TenShop;
                    worksheet.Cells[row + 2, 1].Value = "Mã nhân viên:";
                    worksheet.Cells[row + 2, 2].Value = MaNhanVien;
                    worksheet.Cells[row + 3, 1].Value = "Tên nhân viên:";
                    worksheet.Cells[row + 3, 2].Value = nhanVien.TenNhanVien;
                    worksheet.Cells[row + 4, 1].Value = "Mã hóa đơn:";
                    worksheet.Cells[row + 4, 2].Value = MaHoaDon;
                    worksheet.Cells[row + 5, 1].Value = "Mã khách hàng:";
                    worksheet.Cells[row + 5, 2].Value = tbMaKhachHang.Text;
                    worksheet.Cells[row + 6, 1].Value = "Tên khách hàng:";
                    worksheet.Cells[row + 6, 2].Value = tbTenKhachHang.Text;
                    // ... continue adding the rest of the values in a similar manner

                    // Save Excel package to a file
                    FileInfo fileInfo = new FileInfo(filePath);
                    excelPackage.SaveAs(fileInfo);
                }
                MessageBox.Show("Thanh toán thành công.");
                BLL_QLCF.Instance.ChechOut(MaHoaDon, tbMaKhachHang.Text, ((CBBItems)cbGiamGia.SelectedItem).Text, tbSauGiamGia.Text, tbSoTienKhachTra.Text, tbSoTienThoiLai.Text);
                this.Dispose();
            }
        }

        private void btnPdf_Click(object sender, EventArgs e)
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
            if ((int)Convert.ToInt32(tbDiemTichLuy.Text) < 10 && cbGiamGia.SelectedIndex > 0)
            {
                MessageBox.Show("Khách hàng không đủ điểm để giảm giá", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if ((int)Convert.ToInt32(tbDiemTichLuy.Text) < 20 && cbGiamGia.SelectedIndex > 1)
            {
                MessageBox.Show("Khách hàng không đủ điểm để giảm giá", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (tbSoTienKhachTra.Text == "")
            {
                MessageBox.Show("Vui lòng nhập vào số tiền khách trả", "Chú ý", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show("Bạn có chắc muốn thanh toán không ?", "Thông báo", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                string outputPath = @"E:\BT C#\Coffee-shop(baocao)\Hoá đơn\PDF";
                string pdfFileName = MaHoaDon + ".pdf"; // You need to define MaHoaDon
                string pdfPath = Path.Combine(outputPath, pdfFileName);

                using (FileStream fs = new FileStream(pdfPath, FileMode.Create))
                {
                    Document doc = new Document();
                    PdfWriter writer = PdfWriter.GetInstance(doc, fs);
                    doc.Open();
                    //Hỗ trợ tiếng việt
                    BaseFont bf = BaseFont.CreateFont("c:\\windows\\fonts\\times.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    iTextSharp.text.Font font = new iTextSharp.text.Font(bf, 12);
                    // Add content to the PDF
                    doc.Add(new Paragraph("Shop: " + shop.TenShop));
                    doc.Add(new Paragraph("-----------------------------"));
                    doc.Add(new Paragraph("Mã nhân viên: " + MaNhanVien));
                    doc.Add(new Paragraph("Tên nhân viên: " + nhanVien.TenNhanVien));
                    doc.Add(new Paragraph("Mã hoá đơn: " + MaHoaDon));
                    doc.Add(new Paragraph("Mã khách hàng: " +tbMaKhachHang.Text));
                    doc.Add(new Paragraph("Tên khách hàng: " + tbTenKhachHang.Text));
                    doc.Add(new Paragraph("-----------------------------"));
                    List<Bill> listChiTietHoaDon = BLL_QLCF.Instance.GetListBillByTable(MaBan);
                    foreach (Bill item in listChiTietHoaDon)
                    {
                        doc.Add(new Paragraph("TenMonAn: " + item.TenMonAn));
                        doc.Add(new Paragraph("Số lượng: " + item.SoLuong));
                        doc.Add(new Paragraph("Mã hoá đơn: " + item.ThanhTien));
                    }
                    doc.Add(new Paragraph("-----------------------------"));
                    doc.Add(new Paragraph("Tổng tiền: " + tbTongTien.Text));
                    doc.Add(new Paragraph("Giảm giá: " + cbGiamGia.Text));
                    doc.Add(new Paragraph("Sau giảm giá: " + tbSauGiamGia.Text));
                    doc.Add(new Paragraph("Số tiền khách trả: " + tbSoTienKhachTra.Text));
                    doc.Add(new Paragraph("Số tiền thối lại: " + tbSoTienThoiLai.Text));
                    doc.Add(new Paragraph("Hẹn gặp lại quý khách !"));
                    // Add other information here

                    doc.Close();
                }

                // Perform your other operations here
                BLL_QLCF.Instance.ChechOut(MaHoaDon, tbMaKhachHang.Text, ((CBBItems)cbGiamGia.SelectedItem).Text, tbSauGiamGia.Text, tbSoTienKhachTra.Text, tbSoTienThoiLai.Text);
                this.Dispose();
            }

        }

        private void dgvCustomer_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


        }
    }
}
