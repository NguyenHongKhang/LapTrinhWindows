using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsFormsApp1.DAL;
using WindowsFormsApp1.DTO;

namespace WindowsFormsApp1.BLL
{
    class BLL_QLCF
    {
        private static BLL_QLCF _Instance;
        public static BLL_QLCF Instance
        {
            get

            {
                if (_Instance == null) _Instance = new BLL_QLCF();
                return _Instance;
            }
            private set
            {

            }
        }

        private BLL_QLCF()
        {

        }

        public Account GetAccount(string UserName, string PassWord)
        {
            Account data;
            data = DAL_QLCF.Instance.GetAccount(UserName, PassWord);
            return data;
        }

        public List<Table> GetAllTable()
        {
            List<Table> data = new List<Table>();
            data = DAL_QLCF.Instance.GetAllTable();
            return data;
        }
        public AddressShop GetShop()
        {
            AddressShop data;
            data = DAL_QLCF.Instance.GetShop();
            return data;
        }
        public bool UpdateShop(AddressShop Shop, string NameShop, string TenChuShop, string DiaChiShop, string SoDienThoai)
        {
            if (NameShop == "") throw new InvalidOperationException("Chưa nhập tên shop !!!");
            else if (TenChuShop == "") throw new InvalidOperationException("Chưa nhập tên chủ shop !!!");
            else if (DiaChiShop == "") throw new InvalidOperationException("Chưa nhập địa chỉ shop !!!");
            else if (SoDienThoai == "") throw new InvalidOperationException("Chưa nhập số địne thoại shop !!!");
            else return DAL_QLCF.Instance.UpdateShop(Shop);
        }
        public bool UpdateAccount(Account ac, string MatKhauMoi, string XacNhanMKM)
        {
            if (DAL_QLCF.Instance.GetAccount(ac.TenTaiKhoan, ac.MatKhau) == null)
            { 
                throw new InvalidOperationException("Mật khẩu nhập vào không chính xác!!!");
            }
            if (MatKhauMoi != XacNhanMKM)
            {
                throw new InvalidOperationException("Mật khẩu nhập lại không trùng khớp!!!");
            }
            if (MatKhauMoi == ac.MatKhau)
            {
                throw new InvalidOperationException("Mật khẩu mới trùng với mật khẩu cũ!!!");
            }
            return DAL_QLCF.Instance.UpdateAccount(ac, MatKhauMoi);
        }
        private int AscendingByName(DoAn d1, DoAn d2)
        {
            return d1.TenMonAn.CompareTo(d2.TenMonAn);
        }
        private int DescendingByName(DoAn d1, DoAn d2)
        {
            return d2.TenMonAn.CompareTo(d1.TenMonAn);
        }
        private int AscendingByCategory(Category d1, Category d2)
        {
            return d1.TenLoaiMonAn.CompareTo(d2.TenLoaiMonAn);
        }
        private int DescendingByCategory(Category d1, Category d2)
        {
            return d2.TenLoaiMonAn.CompareTo(d1.TenLoaiMonAn);
        }
        private int AscendingByStaff(NhanVien d1, NhanVien d2)
        {
            return d1.TenNhanVien.CompareTo(d2.TenNhanVien);
        }
        private int DescendingByStaff(NhanVien d1, NhanVien d2)
        {
            return d2.TenNhanVien.CompareTo(d1.TenNhanVien);
        }
        private int AscendingByAccount(Account d1, Account d2)
        {
            return d1.TenTaiKhoan.CompareTo(d2.TenTaiKhoan);
        }
        private int DescendingByAccount(Account d1, Account d2)
        {
            return d2.TenTaiKhoan.CompareTo(d1.TenTaiKhoan);
        }
        public DoAn GetFoodByFoodId(string MaMonAn)
        {
            foreach(DoAn i in DAL_QLCF.Instance.GetAllFood())
            {
                if (i.MaMonAn.Equals(MaMonAn))
                {
                    return i;
                }
            }
            return null;
        }
        public NhanVien GetStaffByStaffId(string MaNhanVien)
        {
            foreach (NhanVien i in DAL_QLCF.Instance.GetAllStaff())
            {
                if (i.MaNhanVien.Equals(MaNhanVien))
                {
                    return i;
                }
            }
            return null;
        }
        public Account GetAccountByAccountName(string TenTaiKhoan)
        {
            foreach (Account i in DAL_QLCF.Instance.GetAllAccount())
            {
                if (i.TenTaiKhoan.Equals(TenTaiKhoan))
                {
                    return i;
                }
            }
            return null;
        }
        public List<DoAn> GetFoodByCategory(string CategoryID)
        {
            List<DoAn> data = new List<DoAn>();
            if (CategoryID.Equals("0"))
            {
                data = DAL_QLCF.Instance.GetAllFood();
            }
            else
            {
                foreach (DoAn i in DAL_QLCF.Instance.GetAllFood())
                {
                    if (i.MaLoaiMonAn.Equals(CategoryID))
                    {
                        data.Add(i);
                    }
                }
            }
            
            return data;
        }
        public List<Account> GetAccountByAccountType(int AccountType)
        {
            List<Account> data = new List<Account>();
            if (AccountType == 2)
            {
                data = DAL_QLCF.Instance.GetAllAccount();
            }
            else
            {
                foreach (Account i in DAL_QLCF.Instance.GetAllAccount())
                {
                    if (i.LoaiTaiKhoan == AccountType)
                    {
                        data.Add(i);
                    }
                }
            }

            return data;
        }
        public List<DoAn> GetListFood(string ValueCategory, string ValueSort, string KeyWord)
        {
            List<DoAn> data = new List<DoAn>();
            List<DoAn> data1 = new List<DoAn>();
            data = GetFoodByCategory(ValueCategory);
            if (ValueSort.Equals("1")) data.Sort(AscendingByName);
            if (ValueSort.Equals("2")) data.Sort(DescendingByName);

            foreach(DoAn i in data)
            {
                if (Fuzzy.Instance.Ratio(i.TenMonAn, KeyWord) >= 0.8) data1.Add(i);
            }
            return data1;
        }
        public List<Category> GetListCategory(string ValueSort, string KeyWord)
        {
            List<Category> data = new List<Category>();
            List<Category> data1 = new List<Category>();
            data = GetAllCategories();
            if (ValueSort.Equals("1")) data.Sort(AscendingByCategory);
            if (ValueSort.Equals("2")) data.Sort(DescendingByCategory);

            foreach (Category i in data)
            {
                if (Fuzzy.Instance.Ratio(i.TenLoaiMonAn, KeyWord) >= 0.8) data1.Add(i);
            }
            return data1;
        }
        public List<NhanVien> GetListStaff(string ValueSort, string KeyWord)
        {
            List<NhanVien> data = new List<NhanVien>();
            List<NhanVien> data1 = new List<NhanVien>();
            data = GetAllStaff();
            if (ValueSort.Equals("1")) data.Sort(AscendingByStaff);
            if (ValueSort.Equals("2")) data.Sort(DescendingByStaff);

            foreach (NhanVien i in data)
            {
                if (Fuzzy.Instance.Ratio(i.TenNhanVien, KeyWord) >= 0.8) data1.Add(i);
            }
            return data1;
        }
        public List<Account> GetListAccount(int ValueCategory, string ValueSort, string KeyWord)
        {
            List<Account> data = new List<Account>();
            List<Account> data1 = new List<Account>();
            data = GetAccountByAccountType(ValueCategory);
            if (ValueSort.Equals("1")) data.Sort(AscendingByAccount);
            if (ValueSort.Equals("2")) data.Sort(DescendingByAccount);

            foreach (Account i in data)
            {
                if (Fuzzy.Instance.Ratio(i.TenTaiKhoan, KeyWord) >= 0.8) data1.Add(i);
            }
            return data1;
        }
        public List<KhachHang> GetListCustomer(string KeyWord)
        {
            List<KhachHang> data = new List<KhachHang>();
            foreach(KhachHang i in DAL_QLCF.Instance.GetAllCustomer())
            {
                if (Fuzzy.Instance.Ratio(i.TenKhachHang, KeyWord) >= 0.8) data.Add(i);
            }
            return data;
        }
        public List<Table> FindTable(string KeyWord)
        {
            List<Table> data = new List<Table>();
            List<Table> allfood = DAL_QLCF.Instance.GetAllTable();
            foreach (Table da in allfood)
            {
                if (Fuzzy.Instance.Ratio(da.TenBan, KeyWord) > 0.8)
                {
                    data.Add(da);
                }
            }
            return data;
        }
        public List<CBBItem> GetCategoryCBBItem()
        {
            List<CBBItem> data = new List<CBBItem>();
            foreach (Category i in DAL_QLCF.Instance.GetAllCategories())
            {
                data.Add(new CBBItem { Text = i.TenLoaiMonAn, Value = i.MaLoaiMonAn });
            }
            return data;
        }
        public List<CBBItems> GetAccountTypeCBBItem()
        {
            List<CBBItems> data = new List<CBBItems>();
            data.Add(new CBBItems { Text = "Nhân viên", Value = 0 });
            data.Add(new CBBItems { Text = "Admin", Value = 1});
            return data;
        }
        public List<Category> GetAllCategories()
        {
            return DAL_QLCF.Instance.GetAllCategories();
        }
        public List<NhanVien> GetAllStaff()
        {
            return DAL_QLCF.Instance.GetAllStaff();
        }
        public List<Account> GetAllAccount()
        {
            return DAL_QLCF.Instance.GetAllAccount();
        }
        public string GetNewFoodId()
        {
            int id = Convert.ToInt32(DAL_QLCF.Instance.GetMaximumFoodId().Substring(2, 3));
            id++;
            string res = "MA" + string.Format("{0:000}", id);
            return res;
        }
        public string GetNewStaffId()
        {
            int id = Convert.ToInt32(DAL_QLCF.Instance.GetMaximumStaffId().Substring(2, 3));
            id++;
            string res = "NV" + string.Format("{0:000}", id);
            return res;
        }
        public bool InsertFood(DoAn da)
        {
            if (da.TenMonAn.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên món ăn!!!");
            }
            if (da.GiaTien <= 0)
            {
                throw new InvalidOperationException("Giá tiền của món ăn không hợp lệ!!!");
            }
            return DAL_QLCF.Instance.InsertFood(da);
        }
        public bool InsertCategory(Category da)
        {
            if (da.TenLoaiMonAn.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên loại món ăn!!!");
            }
            return DAL_QLCF.Instance.InsertCategory(da);
        }
        public bool InsertTable(Table da)
        {
            if (da.TenBan.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên bàn ăn!!!");
            }
            return DAL_QLCF.Instance.InsertTable(da);
        }
        public bool InsertStaff(NhanVien da)
        {
            if (da.TenNhanVien.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên nhân viên!!!");
            }
            if (da.SoDienThoai.Length <= 9)
            {
                throw new InvalidOperationException("Số điện thoại không hợp lệ!!!");
            }
            if (da.DiaChi.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập địa chỉ nhân viên!!!");
            }
            if (da.CMND.Length < 9)
            {
                throw new InvalidOperationException("Số CMND không hợp lệ!!!");
            }
            if (da.QueQuan.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập quê quán nhân viên!!!");
            }
            if (da.CaLamViec.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập ca làm việc!!!");
            }
            return DAL_QLCF.Instance.InsertStaff(da);
        }
        public bool InsertAccount(Account da)
        {
            if (da.TenTaiKhoan.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên tài khoản!!!");
            }
            if (da.MatKhau.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập mật khẩu!!!");
            }
            if (da.TenHienThi.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên hiển thị!!!");
            }
            List<Account> data = new List<Account>();
            List<Account> allfood = DAL_QLCF.Instance.GetAllAccount();
            foreach (Account ac in allfood)
            {
                if (ac.TenTaiKhoan.Equals(da.TenTaiKhoan))
                {
                    data.Add(ac);
                }
            }
           if (data.Count>0)
           {
                throw new InvalidOperationException("Tên tài khoản đã được sử dụng!!!");
           }
            return DAL_QLCF.Instance.InsertAccount(da);
        }
        public bool UpdateFood(DoAn da)
        {
            if (da.TenMonAn.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên món ăn!!!");
            }
            if (da.GiaTien <= 0)
            {
                throw new InvalidOperationException("Giá tiền của món ăn không hợp lệ!!!");
            }
            return DAL_QLCF.Instance.UpdateFood(da);
        }
        public bool UpdateCategory(Category da)
        {
            if (da.TenLoaiMonAn.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên món ăn!!!");
            }
            return DAL_QLCF.Instance.UpdateCategory(da);
        }
        public bool UpdateTable(Table da)
        {
            if (da.TenBan.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên bàn ăn!!!");
            }
            return DAL_QLCF.Instance.UpdateTable(da);
        }
        public bool UpdateStaff(NhanVien da)
        {
            if (da.TenNhanVien.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên nhân viên!!!");
            }
            if (da.SoDienThoai.Length <= 9)
            {
                throw new InvalidOperationException("Số điện thoại không hợp lệ!!!");
            }
            if (da.DiaChi.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập địa chỉ nhân viên!!!");
            }
            if (da.CMND.Length < 9)
            {
                throw new InvalidOperationException("Số CMND không hợp lệ!!!");
            }
            if (da.DiaChi.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập quê quán nhân viên!!!");
            }
            if (da.DiaChi.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập ca làm việc!!!");
            }
            return DAL_QLCF.Instance.UpdateStaff(da);
        }
        public bool UpdateAccount(Account da)
        {
            if (da.MatKhau.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập mật khẩu!!!");
            }
            if (da.TenHienThi.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên hiển thị!!!");
            }
            return DAL_QLCF.Instance.UpdateAccount(da);
        }
        public bool DeleteFood(List<string> ListDelete)
        {
            return DAL_QLCF.Instance.DeleteFood(ListDelete);
        }
        public bool DeleteCategory(List<string> ListDelete)
        {
            return DAL_QLCF.Instance.DeleteCategory(ListDelete);
        }
        public bool DeleteTable(List<string> ListDelete)
        {
            return DAL_QLCF.Instance.DeleteTable(ListDelete);
        }
        public bool DeleteStaff(List<string> ListDelete)
        {
            return DAL_QLCF.Instance.DeleteStaff(ListDelete);
        }
        public bool DeleteAccount(List<string> ListDelete)
        {
            return DAL_QLCF.Instance.DeleteAccount(ListDelete);
        }
        public string GetUncheckBillIDByTableID(string TableID)
        {
            return DAL_QLCF.Instance.GetUncheckBillIDByTableID(TableID);
        }
        public string GetNewBillID()
        {
            int id = Convert.ToInt32(DAL_QLCF.Instance.GetMaximunBillID().Substring(2, 3));
            id++;
            string res = "HD" + string.Format("{0:000}", id);
            return res;
        }
        public void InsertIntoBill(string MaHoaDon, string MaBan, string MaNhanVien)
        {
            DAL_QLCF.Instance.InsertIntoBill(MaHoaDon,MaBan,MaNhanVien);
        }
        public bool InsertBillData(string MaHoaDon, string MaMonAn, int SoLuong)
        {
            return DAL_QLCF.Instance.InsertBillData(MaHoaDon, MaMonAn, SoLuong);
        }
        public List<Bill> GetListBillByTable(string MaBan)
        {
            return DAL_QLCF.Instance.GetListBillByTable(MaBan);
        }
        public void ChechOut(string MaHoaDon, string MaKhachHang, string GiamGia, string TongTien, string SoTienKhachTra, string SoTienThoiLai)
        {
            DAL_QLCF.Instance.ChechOut(MaHoaDon, MaKhachHang, GiamGia, TongTien, SoTienKhachTra, SoTienThoiLai);
        }

        public List<KhachHang> GetAllCustomer()
        {
            return DAL_QLCF.Instance.GetAllCustomer();
        }
        public string GetNewCustomerId()
        {
            int id = Convert.ToInt32(DAL_QLCF.Instance.GetMaximumCustomerId().Substring(2, 3));
            id++;
            string res = "KH" + string.Format("{0:000}", id);
            return res;
        }
        public string GetNewCategoryrId()
        {
            int id = Convert.ToInt32(DAL_QLCF.Instance.GetMaximumCategoryId().Substring(3, 3));
            id++;
            string res = "LMA" + string.Format("{0:000}", id);
            return res;
        }
        public string GetNewTableId()
        {
            int id = Convert.ToInt32(DAL_QLCF.Instance.GetMaximumTableId().Substring(2, 3));
            id++;
            string res = "MB" + string.Format("{0:000}", id);
            return res;
        }
        public bool InsertCustomer(KhachHang kh)
        {
            if (kh.TenKhachHang.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên khách hàng!!!");
            }
            if (kh.SoDienThoai.Length <= 9)
            {
                throw new InvalidOperationException("Số điện thoại của khách hàng không hợp lệ!!!");
            }
            return DAL_QLCF.Instance.InsertCustomer(kh);
        }
        public List<KhachHang> FindCustomer(string KeyWord)
        {
            List<KhachHang> data = new List<KhachHang>();
            List<KhachHang> allcustomer = DAL_QLCF.Instance.GetAllCustomer();
            foreach (KhachHang kh in allcustomer)
            {
                if (Fuzzy.Instance.Ratio(kh.TenKhachHang, KeyWord) > 0.8)
                {
                    data.Add(kh);
                }
            }
            return data;
        }
        public bool UpdateCustomer(KhachHang kh)
        {
            if (kh.TenKhachHang.Equals(""))
            {
                throw new InvalidOperationException("Chưa nhập tên khách hàng!!!");
            }
            if (kh.SoDienThoai.Length <= 9)
            {
                throw new InvalidOperationException("Số điện thoại của khách hàng không hợp lệ!!!");
            }
            if (kh.DiemTichLuy < 0)
            {
                throw new InvalidOperationException("Số điểm tích lũy không hợp lệ!!!");
            }
            return DAL_QLCF.Instance.UpdateCustomer(kh);
        }
        public bool DeleteCustomer(List<string> ListDelete)
        {
            return DAL_QLCF.Instance.DeleteCustomer(ListDelete);
        }
        
        public bool MoveTable(Table src, Table des)
        {
            if (src.MaBan.Equals(des.MaBan))
            {
                throw new InvalidOperationException("Trùng bàn!!!");
            }
            if(src.TrangThai.Equals("Trống") && des.TrangThai.Equals("Trống"))
            {
                throw new InvalidOperationException("Hai bàn được chuyển đều là bàn trống!!!");
            }
            return DAL_QLCF.Instance.MoveTable(src, des);
        }
    
        public bool MergeTable(Table src, Table des)
        {
            if (src.MaBan.Equals(des.MaBan))
            {
                throw new InvalidOperationException("Trùng bàn!!!");
            }
            if (src.TrangThai.Equals("Trống") || des.TrangThai.Equals("Trống"))
            {
                throw new InvalidOperationException("Bàn được gộp phải là bàn đang sử dụng!!!");
            }
            return DAL_QLCF.Instance.MergeTable(src, des);
        }
        
        public DataTable TDT_GetRevenueTable(string TimeFormat, string TimeMax, string TimeMin)
        {
            DataTable dt = DAL_QLCF.Instance.TDT_GetRevenueTable(TimeFormat, TimeMax, TimeMin);
            return dt;
        }

        public DataTable NV_GetRevenueTable(string Shift, string TimeMax, string TimeMin)
        {
            DataTable dt = DAL_QLCF.Instance.NV_GetRevenueTable(Shift, TimeMax, TimeMin);
            return dt;
        }

        public DataTable MA_GetRevenueTable(string Category, string TimeMax, string TimeMin)
        {
            DataTable dt = DAL_QLCF.Instance.MA_GetRevenueTable(Category, TimeMax, TimeMin);
            return dt;
        }

        public DataTable LoadListBillByDate(DateTime checkIn, DateTime checkOut)
        {
            return DAL_QLCF.Instance.GetListBillsByDate(checkIn, checkOut);
        }

    }
}
