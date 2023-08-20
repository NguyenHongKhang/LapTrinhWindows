using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.DTO
{
    public class Account
    {
        public Account(string TenTaiKhoan, string MatKhau, string TenHienThi, int LoaiTaiKhoan, string MaNhanVien)
        {
            this.TenTaiKhoan = TenTaiKhoan;
            this.MatKhau = MatKhau;
            this.TenHienThi = TenHienThi;
            this.LoaiTaiKhoan = LoaiTaiKhoan;
            this.MaNhanVien = MaNhanVien;
        }
        public Account(DataRow row)
        {
            this.TenTaiKhoan = row["TenTaiKhoan"].ToString();
            this.MatKhau = row["MatKhau"].ToString();
            this.TenHienThi = row["TenHienThi"].ToString();
            this.LoaiTaiKhoan = (int)row["LoaiTaiKhoan"];
            this.MaNhanVien = row["MaNhanVien"].ToString();
        }
        private string _TenTaiKhoan;

        public string TenTaiKhoan 
        {
            get { return _TenTaiKhoan; }
            set { _TenTaiKhoan = value; }
        }

        private string _MatKhau;

        public string MatKhau
        {
            get { return _MatKhau; }
            set { _MatKhau = value; }
        }
        private string _TenHienThi;

        public string TenHienThi
        {
            get { return _TenHienThi; }
            set { _TenHienThi = value; }
        }
        private int _LoaiTaiKhoan;
        public int LoaiTaiKhoan
        {
            get { return _LoaiTaiKhoan; }
            set { _LoaiTaiKhoan = value; }
        }
        private string _MaNhanVien;
        public string MaNhanVien
        {
            get { return _MaNhanVien; }
            set { _MaNhanVien = value; }
        }
    }
}
