using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.DTO
{
    class ChiTietHoaDon
    {
        public ChiTietHoaDon(int MaHoaDon, string TenKhachHang, string TenNhanVien, string MaMonAn, int SoLuong)
        {
            this.MaHoaDon = MaHoaDon;
            this.TenKhachHang = TenKhachHang;
            this.TenNhanVien = TenNhanVien;
            this.MaMonAn = MaMonAn;
            this.SoLuong = SoLuong;

        }
        public ChiTietHoaDon(DataRow row)
        {
            this.MaHoaDon = (int)row["MaHoaDon"];
            this.TenKhachHang = (string)row["TenKhachHang"];
            this.TenNhanVien = (string)row["TenNhanVien"];
            this.MaMonAn = row["MaMonAn"].ToString();
            this.SoLuong = (int)row["SoLuong"];
        }
        private int _MaHoaDon;
        public int MaHoaDon
        {
            get { return _MaHoaDon; }
            set { _MaHoaDon = value; }
        }

        private string _TenKhachHang;
        public string TenKhachHang
        {
            get { return _TenKhachHang; }
            set { _TenKhachHang = value; }
        }
        private string _TenNhanVien;
        public string TenNhanVien
        {
            get { return _TenNhanVien; }
            set { _TenNhanVien = value; }
        }
        private string _MaMonAn;
        public string MaMonAn
        {
            get { return _MaMonAn; }
            set { _MaMonAn = value; }
        }
        private int _SoLuong;
        public int SoLuong
        {
            get { return _SoLuong; }
            set { _SoLuong = value; }
        }


    }
}
