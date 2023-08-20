using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.DTO
{
    public class KhachHang
    {
        public KhachHang(string TenKhachHang, string MaKhachHang, string SoDienThoai, int SoBillsDaThanhToan, int DiemTichLuy)
        {
            this.TenKhachHang = TenKhachHang;
            this.MaKhachHang = MaKhachHang;
            this.SoDienThoai = SoDienThoai;

            this.SoBillsDaThanhToan = SoBillsDaThanhToan;
            this.DiemTichLuy = DiemTichLuy;
        }
        public KhachHang(DataRow row)
        {
            this.TenKhachHang = row["TenKhachHang"].ToString();
            this.MaKhachHang = row["MaKhachHang"].ToString();
            this.SoDienThoai = row["SoDienThoai"].ToString();
            this.SoBillsDaThanhToan = (int)row["SoBillsDaThanhToan"];
            this.DiemTichLuy = (int)row["DiemTichLuy"];
        }
        private string _TenKhachHang;

        public string TenKhachHang 
        {
            get { return _TenKhachHang; }
            set { _TenKhachHang = value; }
        }
        private string _MaKhachHang;

        public string MaKhachHang
        {
            get { return _MaKhachHang; }
            set { _MaKhachHang = value; }
        }
        private string _SoDienThoai;

        public string SoDienThoai
        {
            get { return _SoDienThoai; }
            set { _SoDienThoai = value; }
        }
        private int _SoBillsDaThanhToan;

        public int SoBillsDaThanhToan
        {
            get { return _SoBillsDaThanhToan; }
            set { _SoBillsDaThanhToan = value; }
        }
        private int _DiemTichLuy;

        public int DiemTichLuy
        {
            get { return _DiemTichLuy; }
            set { _DiemTichLuy = value; }
        }
    }
}
