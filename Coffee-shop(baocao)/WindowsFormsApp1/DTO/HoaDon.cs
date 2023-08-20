using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.DTO
{
    public class HoaDon
    {
        public HoaDon(string MaHoaDon,string MaBan, DateTime? ThoiGianVao, DateTime? ThoiGianRa, string TrangThai, string MaKhachHang, string MaNhanVien, float TongTien, string GiamGia, float SoTienKhachTra, float SoTienThoiLai)
        {
            this.MaHoaDon = MaHoaDon;
            this.MaBan = MaBan;
            this.ThoiGianVao = ThoiGianVao;
            this.ThoiGianRa = ThoiGianRa;

            this.TrangThai = TrangThai;
            this.MaKhachHang = MaKhachHang;
            this.TongTien = TongTien;
            this.GiamGia = GiamGia;
            this.SoTienKhachTra = SoTienKhachTra;
            this.SoTienThoiLai = SoTienThoiLai;
        }
        public HoaDon(DataRow row)
        {
            this.MaHoaDon = row["MaHoaDon"].ToString();
            this.MaBan = row["MaBan"].ToString();
            this.ThoiGianVao = (DateTime?)ThoiGianVao;
            var ThoiGianRaTemp = row["ThoiGianRa"];
            if (ThoiGianRaTemp.ToString()!="")
                this.ThoiGianRa = (DateTime?)ThoiGianRaTemp;
            this.TrangThai = (string)row["TrangThai"];
            this.MaKhachHang = row["MaKhachHang"].ToString();
            var TongTienTemp = row["TongTien"];
            if (TongTienTemp.ToString() != "")
                this.TongTien = (float)Convert.ToDouble(row["TongTien"].ToString());
            this.GiamGia = row["GiamGia"].ToString();
            var SoTienKhachTraTemp = row["SoTienKhachTra"];
            if (SoTienKhachTraTemp.ToString() != "")
                this.SoTienKhachTra = (float)Convert.ToDouble(row["SoTienKhachTra"].ToString());
            var SoTienThoiLaiTemp = row["SoTienKhachTra"];
            if (SoTienThoiLaiTemp.ToString() != "")
                this.SoTienThoiLai = (float)Convert.ToDouble(row["SoTienThoiLai"].ToString());
        }
        private string _TrangThai;
        public string TrangThai
        {
            get { return _TrangThai; }
            set { _TrangThai = value; }
        }

        private DateTime? _ThoiGianVao;
        public DateTime? ThoiGianVao
        {
            get { return _ThoiGianVao; }
            set { _ThoiGianVao = value; }
        }
        private DateTime? _ThoiGianRa;
        public DateTime? ThoiGianRa
        {
            get { return _ThoiGianRa; }
            set { _ThoiGianRa = value; }
        }

        private string _MaHoaDon;

        public string MaHoaDon 
        {
            get { return _MaHoaDon; }
            set { _MaHoaDon = value; } 
        }
        private string _MaBan;

        public string MaBan
        {
            get { return _MaBan; }
            set { _MaBan = value; }
        }
        private string _MaKhachHang;

        public string MaKhachHang
        {
            get { return _MaKhachHang; }
            set { _MaKhachHang = value; }
        }
        private string _MaNhanVien;

        public string MaNhanVien
        {
            get { return _MaNhanVien; }
            set { _MaNhanVien = value; }
        }
        private float _TongTien;

        public float TongTien
        {
            get { return _TongTien; }
            set { _TongTien = value; }
        }
        private string _GiamGia;

        public string GiamGia
        {
            get { return _GiamGia; }
            set { _GiamGia = value; }
        }
        private float _SoTienKhachTra;

        public float SoTienKhachTra
        {
            get { return _SoTienKhachTra; }
            set { _SoTienKhachTra = value; }
        }
        private float _SoTienThoiLai;

        public float SoTienThoiLai
        {
            get { return _SoTienThoiLai; }
            set { _SoTienThoiLai = value; }
        }
    }
}
