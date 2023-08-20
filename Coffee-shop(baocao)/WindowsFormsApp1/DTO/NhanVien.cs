using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.DTO
{
    public class NhanVien
    {
        public NhanVien (string TenNhanVien, string MaNhanVien, string SoDienThoai, string DiaChi, string CMND, string QueQuan, string CaLamViec, Byte[] HinhAnh)
        {
            this.TenNhanVien = TenNhanVien;
            this.MaNhanVien = MaNhanVien;
            this.SoDienThoai = SoDienThoai;
            this.DiaChi = DiaChi;
            this.CMND = CMND;

            this.QueQuan = QueQuan;
            this.CaLamViec = CaLamViec;
            this.HinhAnh = HinhAnh;
        }
        public NhanVien (DataRow row)
        {
            this.TenNhanVien = row["TenNhanVien"].ToString();
            this.MaNhanVien = row["MaNhanVien"].ToString();
            this.SoDienThoai = row["SoDienThoai"].ToString();
            this.DiaChi = row["DiaChi"].ToString();
            this.CMND = row["CMND"].ToString();
            this.QueQuan = row["QueQuan"].ToString();
            this.CaLamViec = row["CaLamViec"].ToString();
            if (row["HinhAnh"].GetType() != typeof(System.DBNull))
            {
                byte[] temp = (byte[])row["HinhAnh"];
                if (temp.Length != 0) this.HinhAnh = temp;
                else this.HinhAnh = null;
            }
            else this.HinhAnh = null;
        }
        private string _TenNhanVien;
        public string TenNhanVien 
        {
            get { return _TenNhanVien; }
            set { _TenNhanVien = value; }
        }
        private string _MaNhanVien;
        public string MaNhanVien
        {
            get { return _MaNhanVien; }
            set { _MaNhanVien = value; }
        }
        private string _SoDienThoai;
        public string SoDienThoai
        {
            get { return _SoDienThoai; }
            set { _SoDienThoai = value; }
        }
        private string _DiaChi;
        public string DiaChi
        {
            get { return _DiaChi; }
            set { _DiaChi = value; }
        }
        private string _CMND;
        public string CMND
        {
            get { return _CMND; }
            set { _CMND = value; }
        }
        private string _QueQuan;
        public string QueQuan
        {
            get { return _QueQuan; }
            set { _QueQuan = value; }
        }
        private string _CaLamViec;
        public string CaLamViec
        {
            get { return _CaLamViec; }
            set { _CaLamViec = value; }
        }
        private Byte[] _HinhAnh;
        public Byte[] HinhAnh
        {
            get { return _HinhAnh; }
            set { _HinhAnh = value; }
        }
    }
}
