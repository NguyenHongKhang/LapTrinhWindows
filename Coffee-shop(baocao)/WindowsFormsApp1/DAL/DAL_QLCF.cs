using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsFormsApp1.DAO;
using WindowsFormsApp1.DTO;

namespace WindowsFormsApp1.DAL
{

    class DAL_QLCF
    {
        private static DAL_QLCF _Instance;
        public static DAL_QLCF Instance
        {
            get
            {
                if (_Instance == null) _Instance = new DAL_QLCF();
                return _Instance;
            }
            private set { }
        }

        private DAL_QLCF()
        {

        }
        public Account GetAccount(string TenTaiKhoan, string MatKhau)
        {
            Account data = null;
            string query = "USP_Login @TenTaiKhoan , @MatKhau";
            DataTable result = DataProvider.Instance.ExecuteQuery(query, new object[] { TenTaiKhoan, MatKhau });
            if (result.Rows.Count > 0)
            {
                data = new Account(result.Rows[0]);
            }
            return data;
        }

        public List<Table> GetAllTable()
        {
            List<Table> data = new List<Table>();

            DataTable dt = DataProvider.Instance.ExecuteQuery("USP_GetTableList");

            foreach (DataRow item in dt.Rows)
            {
                data.Add(new Table(item));
            }

            return data;
        }

        public bool UpdateAccount(Account ac, string MatKhauMoi)
        {
            int result = DataProvider.Instance.ExecuteNonQuery("USP_UpdateAccount @TenTaiKhoan , @TenHienThi , @MatKhau , @MatKhauMoi", new object[] { ac.TenTaiKhoan, ac.TenHienThi, ac.MatKhau, MatKhauMoi });
            return result > 0;
        }

        public List<DoAn> GetAllFood()
        {
            List<DoAn> list = new List<DoAn>();
            string query = "USP_GetListFood";
            DataTable data = DataProvider.Instance.ExecuteQuery(query);
            foreach (DataRow i in data.Rows)
            {
                DoAn doan = new DoAn(i);
                list.Add(doan);
            }
            return list;
        }

        public List<Category> GetAllCategories()
        {
            List<Category> data = new List<Category>();
            string query = "select * from danhmuc where HieuLuc = 1";
            DataTable dt = DataProvider.Instance.ExecuteQuery(query);
            foreach (DataRow i in dt.Rows)
            {
                Category ca = new Category(i);
                data.Add(ca);
            }
            return data;
        }
        public List<NhanVien> GetAllStaff()
        {
            List<NhanVien> data = new List<NhanVien>();
            string query = "select * from NhanVien where HieuLuc = 1";
            DataTable dt = DataProvider.Instance.ExecuteQuery(query);
            foreach (DataRow i in dt.Rows)
            {
                NhanVien ca = new NhanVien(i);
                data.Add(ca);
            }
            return data;
        }
        public List<Account> GetAllAccount()
        {
            List<Account> data = new List<Account>();
            string query = "select * from DangNhap where HieuLuc = 1";
            DataTable dt = DataProvider.Instance.ExecuteQuery(query);
            foreach (DataRow i in dt.Rows)
            {
                Account ca = new Account(i);
                data.Add(ca);
            }
            return data;
        }
        public string GetMaximumFoodId()
        {
            string query = "select max(mamonan) from DoAn";
            string MaximumFoodId = "";
            try
            {
                MaximumFoodId = (string) DataProvider.Instance.ExecuteScalarQuery(query);
            }
            catch
            {
                MaximumFoodId = "MA000";
            }
            return MaximumFoodId;
        }
        public string GetMaximumStaffId()
        {
            string query = "select max(manhanvien) from NhanVien";
            string MaximumStaffId = "";
            try
            {
                MaximumStaffId = (string)DataProvider.Instance.ExecuteScalarQuery(query);
            }
            catch
            {
                MaximumStaffId = "NV000";
            }
            return MaximumStaffId;
        }
        public AddressShop GetShop()
        {
            AddressShop data = null;
            string query = "select TenShop , TenChuShop , DiaChiShop ,SoDienThoai from dbo.Shop ";
            DataTable result = DataProvider.Instance.ExecuteQuery(query);
            if (result.Rows.Count > 0)
            {
                data = new AddressShop(result.Rows[0]);
            }
            return data;
        }
        public bool InsertFood(DoAn da)
        {
            string query = "exec USP_InsertFood @TenMonAn , @MaMonAn , @GiaTien , @MaLoaiMonAn , @TinhTrangMonAn , @HinhAnh";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { da.TenMonAn, da.MaMonAn, da.GiaTien, da.MaLoaiMonAn, da.TinhTrangMonAn, da.HinhAnh });
            return result > 0;
        }
        public bool InsertCategory(Category da)
        {
            string query = "exec USP_InsertCategory @TenLoaiMonAn , @MaLoaiMonAn";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { da.TenLoaiMonAn, da.MaLoaiMonAn});
            return result > 0;
        }
        public bool InsertTable(Table da)
        {
            string query = "exec USP_InsertTable @TenBan , @MaBan , @TrangThai";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { da.TenBan, da.MaBan, da.TrangThai });
            return result > 0;
        }
        public bool InsertStaff(NhanVien da)
        {
            string query = "exec USP_InsertStaff @TenNhanVien , @MaNhanVien , @SoDienThoai , @DiaChi , @CMND , @QueQuan , @CaLamViec , @HinhAnh";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { da.TenNhanVien, da.MaNhanVien, da.SoDienThoai, da.DiaChi, da.CMND, da.QueQuan, da.CaLamViec, da.HinhAnh });
            return result > 0;
        }
        public bool InsertAccount(Account da)
        {
            string query = "exec USP_InsertAccount @TenTaiKhoan , @MatKhau , @TenHienThi , @LoaiTaiKhoan , @MaNhanVien";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { da.TenTaiKhoan, da.MatKhau, da.TenHienThi, da.LoaiTaiKhoan, da.MaNhanVien});
            return result > 0;
        }
        public bool UpdateFood(DoAn da)
        {
            string query = "exec USP_UpdateFood @TenMonAn , @MaMonAn , @GiaTien , @MaLoaiMonAn , @TinhTrangMonAn , @HinhAnh";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { da.TenMonAn, da.MaMonAn, da.GiaTien, da.MaLoaiMonAn, da.TinhTrangMonAn, da.HinhAnh });
            return result > 0;
        }
        public bool UpdateCategory(Category da)
        {
            string query = "exec USP_UpdateCategory @TenLoaiMonAn , @MaLoaiMonAn";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { da.TenLoaiMonAn, da.MaLoaiMonAn});
            return result > 0;
        }
        public bool UpdateTable(Table da)
        {
            string query = "exec USP_UpdateTable @TenBan , @MaBan";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { da.TenBan, da.MaBan });
            return result > 0;
        }
        public bool UpdateStaff(NhanVien da)
        {
            string query = "exec USP_UpdateStaff @TenNhanVien , @MaNhanVien , @SoDienThoai , @DiaChi , @CMND , @QueQuan , @CaLamViec , @HinhAnh";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { da.TenNhanVien, da.MaNhanVien, da.SoDienThoai, da.DiaChi, da.CMND, da.QueQuan, da.CaLamViec, da.HinhAnh });
            return result > 0;
        }
        public bool UpdateAccount(Account da)
        {
            string query = "exec USP_UpdateAccountAdmin @TenTaiKhoan , @MatKhau , @TenHienThi , @LoaiTaiKhoan , @MaNhanVien";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { da.TenTaiKhoan, da.MatKhau, da.TenHienThi, da.LoaiTaiKhoan, da.MaNhanVien });
            return result > 0;
        }
        public bool UpdateShop(AddressShop Shop)
        {
            int result = DataProvider.Instance.ExecuteNonQuery("USP_UpdateShop @TenShop , @TenChuShop , @DiaChiShop , @SoDienThoai ", new object[] { Shop.TenShop, Shop.TenChuShop, Shop.DiaChiShop, Shop.SoDienThoai });
            return result > 0;
        }
        public bool DeleteFood(List<string> ListDelete)
        {
            string query = "USP_DeleteFood @MaMonAn";
            List<object[]> ListParameter = new List<object[]>();
            foreach(string i in ListDelete)
            {
                ListParameter.Add(new object[] {i});
            }
            int result = DataProvider.Instance.ExecuteNonQueryTransaction(query, ListParameter);
            return result > 0;
        }
        public bool DeleteCategory(List<string> ListDelete)
        {
            string query = "USP_DeleteCategory @MaLoaiMonAn";
            List<object[]> ListParameter = new List<object[]>();
            foreach (string i in ListDelete)
            {
                ListParameter.Add(new object[] { i });
            }
            int result = DataProvider.Instance.ExecuteNonQueryTransaction(query, ListParameter);
            return result > 0;
        }
        public bool DeleteTable(List<string> ListDelete)
        {
            string query = "USP_DeleteTable @MaBan";
            List<object[]> ListParameter = new List<object[]>();
            foreach (string i in ListDelete)
            {
                ListParameter.Add(new object[] { i });
            }
            int result = DataProvider.Instance.ExecuteNonQueryTransaction(query, ListParameter);
            return result > 0;
        }
        public bool DeleteStaff(List<string> ListDelete)
        {
            string query = "USP_DeleteStaff @MaNhanVien";
            List<object[]> ListParameter = new List<object[]>();
            foreach (string i in ListDelete)
            {
                ListParameter.Add(new object[] { i });
            }
            int result = DataProvider.Instance.ExecuteNonQueryTransaction(query, ListParameter);
            return result > 0;
        }
        public bool DeleteAccount(List<string> ListDelete)
        {
            string query = "USP_DeleteAccount @TenTaiKhoan";
            List<object[]> ListParameter = new List<object[]>();
            foreach (string i in ListDelete)
            {
                ListParameter.Add(new object[] { i });
            }
            int result = DataProvider.Instance.ExecuteNonQueryTransaction(query, ListParameter);
            return result > 0;
        }
        public string GetUncheckBillIDByTableID(string TableID)
        {
            DataTable data = DataProvider.Instance.ExecuteQuery("Select * from HoaDon where MaBan = N'" + TableID + "' and TrangThai = N'Chưa thanh toán'");
            if (data.Rows.Count > 0)
            {
                HoaDon bill = new HoaDon(data.Rows[0]);
                return bill.MaHoaDon;
            }
            return "-1";
        }

        public string GetMaximunBillID()
        {
            try
            {
                return (string) DataProvider.Instance.ExecuteScalarQuery("select max(MaHoaDon) from HoaDon");
            }
            catch
            {
                return "HD000";
            }
        }

        public void InsertIntoBill(string MaHoaDon, string MaBan, string MaNhanVien)
        {
            DataProvider.Instance.ExecuteNonQuery("exec USP_InsertHoaDon @MaHoaDon , @MaBan , @MaNhanVien ", new object[] { MaHoaDon, MaBan, MaNhanVien });
        }

        public bool InsertBillData(string MaHoaDon, string MaMonAn, int SoLuong)
        {
            int check = DataProvider.Instance.ExecuteNonQuery("exec USP_InsertChiTietHoaDon @MaHoaDon , @MaMonAn , @SoLuong", new object[] { MaHoaDon, MaMonAn, SoLuong });
            return check > 0;
        }

        public List<Bill> GetListBillByTable(string MaBan)
        {
            List<Bill> listMenu = new List<Bill>();

            DataTable data = DataProvider.Instance.ExecuteQuery("Select d.TenMonAn, c.SoLuong, d.GiaTien, d.GiaTien*c.SoLuong as ThanhTien from ChiTietHoaDon as c, HoaDon as h, DoAn as d where c.MaHoaDon = h.MaHoaDon and c.MaMonAn = d.MaMonAn and h.MaBan = N'" + MaBan + "' and h.TrangThai = N'Chưa thanh toán'");

            foreach (DataRow item in data.Rows)
            {
                Bill menu = new Bill(item);
                listMenu.Add(menu);
            }
            return listMenu;
        }

        public void ChechOut(string MaHoaDon, string MaKhachHang, string GiamGia, string TongTien, string SoTienKhachTra, string SoTienThoiLai)
        {
            DataProvider.Instance.ExecuteNonQuery("exec USP_CheckOut @MaHoaDon , @MaKhachHang , @GiamGia , @TongTien , @SoTienKhachTra , @SoTienThoiLai ", new object[] { MaHoaDon, MaKhachHang, GiamGia, TongTien, SoTienKhachTra, SoTienThoiLai });
        }

        public List<KhachHang> GetAllCustomer()
        {
            List<KhachHang> list = new List<KhachHang>();
            string query = "Select * from KhachHang where HieuLuc = 1";
            DataTable data = DataProvider.Instance.ExecuteQuery(query);
            foreach (DataRow i in data.Rows)
            {
                KhachHang customer = new KhachHang(i);
                list.Add(customer);
            }
            return list;
        }
        public DataTable GetListBillsByDate(DateTime checkIn, DateTime checkOut)
        {
            return DataProvider.Instance.ExecuteQuery("exec USP_GetListBillByDate @checkIn , @checkOut", new object[] { checkIn, checkOut });
        }
        public List<KhachHang> GetListCustomerByName(string TenKhachHang)
        {
            List<KhachHang> list = new List<KhachHang>();
            string query = "select * from khachhang where HieuLuc = 1 and tenkhachhang like N'Nguyễn%'";
            DataTable data = DataProvider.Instance.ExecuteQuery(query);
            foreach (DataRow i in data.Rows)
            {
                KhachHang customer = new KhachHang(i);
                list.Add(customer);
            }
            return list;
        }
        public void TableCancel(string MaHoaDon, string MaBan)
        {
            DataProvider.Instance.ExecuteNonQuery("Delete from dbo.ChiTietHoaDon where MaHoaDon = '" + MaHoaDon + "'");
            DataProvider.Instance.ExecuteNonQuery("Delete from dbo.HoaDon where MaHoaDon ='" + MaHoaDon + "'");
            DataProvider.Instance.ExecuteNonQuery("update dbo.BanAn set TrangThai = N'Trống' where MaBan ='" + MaBan + "'");
        }
        public string GetMaximumCustomerId()
        {
            string query = "select max(makhachhang) from KhachHang";
            string MaximumCustomerId = "";
            try
            {
                MaximumCustomerId = (string)DataProvider.Instance.ExecuteScalarQuery(query);
            }
            catch
            {
                MaximumCustomerId = "KH000";
            }
            return MaximumCustomerId;
        }
        public string GetMaximumCategoryId()
        {
            string query = "select max(maloaimonan) from DanhMuc";
            string MaximumCategoryrId = "";
            try
            {
                MaximumCategoryrId = (string)DataProvider.Instance.ExecuteScalarQuery(query);
            }
            catch
            {
                MaximumCategoryrId = "LMA000";
            }
            return MaximumCategoryrId;
        }
        public string GetMaximumTableId()
        {
            string query = "select max(maban) from BanAn";
            string MaximumTableId = "";
            try
            {
                MaximumTableId = (string)DataProvider.Instance.ExecuteScalarQuery(query);
            }
            catch
            {
                MaximumTableId = "MB000";
            }
            return MaximumTableId;
        }
        public bool InsertCustomer(KhachHang kh)
        {
            string query = "exec USP_InsertCustomer @TenKhachHang , @MaKhachHang , @SoDienThoai , @SoBillsDaThanhToan , @DiemTichLuy ";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { kh.TenKhachHang, kh.MaKhachHang, kh.SoDienThoai, kh.SoBillsDaThanhToan, kh.DiemTichLuy});
            return result > 0;
        }
        public bool UpdateCustomer(KhachHang kh)
        {
            string query = "exec USP_UpdateCustomer @TenKhachHang , @MaKhachHang , @SoDienThoai , @SoBillsDaThanhToan , @DiemTichLuy ";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { kh.TenKhachHang, kh.MaKhachHang, kh.SoDienThoai, kh.SoBillsDaThanhToan, kh.DiemTichLuy});
            return result > 0;
        }
        public bool DeleteCustomer(List<string> ListDelete)
        {
            string query = "USP_DeleteCustomer @MaKhachHang";
            List<object[]> ListParameter = new List<object[]>();
            foreach (string i in ListDelete)
            {
                ListParameter.Add(new object[] { i });
            }
            int result = DataProvider.Instance.ExecuteNonQueryTransaction(query, ListParameter);
            return result > 0;
        }
    
        public bool MoveTable(Table src, Table des)
        {
            string query = "USP_MoveTable @MaBan1 , @MaBan2";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { src.MaBan, des.MaBan });
            return result > 0;
        }
        
        public bool MergeTable(Table src, Table des)
        {
            string query = "USP_MergeTable @MaBan1 , @MaBan2";
            int result = DataProvider.Instance.ExecuteNonQuery(query, new object[] { src.MaBan, des.MaBan });
            return result > 0;
        }

        public DataTable TDT_GetRevenueTable(string TimeFormat, string TimeMax, string TimeMin)
        {
            string query = "select FORMAT(ThoiGianRa,'" + TimeFormat + "') AS ThoiGian ,sum(TongTien) AS TongTien from dbo.HoaDon  where ThoiGianRa >= '" + TimeMin + "' AND ThoiGianRa <= '" + TimeMax + "' group by FORMAT(ThoiGianRa,'" + TimeFormat + "') ";
            DataTable dt = DataProvider.Instance.ExecuteQuery(query);
            return dt;
        }

        public DataTable NV_GetRevenueTable(string Shift, string TimeMax, string TimeMin)
        {
            string query = "";
            if(Shift.Equals(""))
            {
                query = "select TenNhanVien, Sum(TongTien) as 'DoanhThu' from dbo.HoaDon inner join dbo.NhanVien on dbo.HoaDon.MaNhanVien = dbo.NhanVien.MaNhanVien where TongTien IS NOT NULL and ThoiGianRa >='" + TimeMin + "' and ThoiGianRa <= '" + TimeMax + "' group by TenNhanVien";
            }
            else
            {
                query = "select TenNhanVien, Sum(TongTien) as 'DoanhThu' from dbo.HoaDon inner join dbo.NhanVien on dbo.HoaDon.MaNhanVien = dbo.NhanVien.MaNhanVien where TongTien IS NOT NULL and CaLamViec = N'" + Shift + "' and ThoiGianRa >='" + TimeMin + "' and ThoiGianRa <= '" + TimeMax + "' group by TenNhanVien";
            }
            DataTable dt = DataProvider.Instance.ExecuteQuery(query);
            return dt;
        }

        public DataTable MA_GetRevenueTable(string Category, string TimeMax, string TimeMin)
        {
            string query = "";
            if (Category.Equals(""))
            {
                query = "select DoAn.TenMonAn, sum(ChiTietHoaDon.ThanhTien) as 'DoanhThu' from DoAn inner join ChiTietHoaDon on DoAn.MaMonAn = ChiTietHoaDon.MaMonAn inner join HoaDon on ChiTietHoaDon.MaHoaDon = HoaDon.MaHoaDon where TongTien IS NOT NULL and ThoiGianRa >='" + TimeMin + "' and ThoiGianRa <= '" + TimeMax + "' group by DoAn.TenMonAn";
            }
            else
            {
                query = "select DoAn.TenMonAn, sum(ChiTietHoaDon.ThanhTien) as 'DoanhThu' from DoAn inner join ChiTietHoaDon on DoAn.MaMonAn = ChiTietHoaDon.MaMonAn inner join HoaDon on ChiTietHoaDon.MaHoaDon = HoaDon.MaHoaDon where TongTien IS NOT NULL and DoAn.MaLoaiMonAn = N'" + Category + "' and ThoiGianRa >='" + TimeMin + "' and ThoiGianRa <= '" + TimeMax + "' group by DoAn.TenMonAn";
            }
            DataTable dt = DataProvider.Instance.ExecuteQuery(query);
            return dt;
        }

    }
}
