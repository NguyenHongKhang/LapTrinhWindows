using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.DTO
{
    public class DoAn
    {
        public DoAn(string TenMonAn, string MaMonAn, float GiaTien, string MaLoaiMonAn, bool TinhTrangMonAn, Byte[] HinhAnh)
        {
            this.TenMonAn = TenMonAn;
            this.MaMonAn = MaMonAn;

            this.GiaTien = GiaTien;
            this.MaLoaiMonAn = MaLoaiMonAn;
            this.TinhTrangMonAn = TinhTrangMonAn;
            this.HinhAnh = HinhAnh;
        }
        public DoAn(DataRow row)
        {
            this.TenMonAn = row["TenMonAn"].ToString();
            this.MaMonAn = row["MaMonAn"].ToString();
            this.GiaTien = (float)Convert.ToDouble(row["GiaTien"].ToString());
            this.MaLoaiMonAn = row["MaLoaiMonAn"].ToString();
            this.TinhTrangMonAn = Convert.ToBoolean(row["TinhTrangMonAn"]);
            if (row["HinhAnh"].GetType() != typeof(System.DBNull))
            {
                byte [] temp = (byte[])row["HinhAnh"];
                if (temp.Length != 0) this.HinhAnh = temp;
                else this.HinhAnh = null;
            }
            else this.HinhAnh = null;
        }
        private string _TenMonAn;

        public string TenMonAn 
        {
            get { return _TenMonAn; }
            set { _TenMonAn = value; }
        }
        private string _MaMonAn;
        public string MaMonAn
        {
            get { return _MaMonAn; }
            set { _MaMonAn = value; }
        }
        private float _GiaTien;
        public float GiaTien
        {
            get { return _GiaTien; }
            set { _GiaTien = value; }
        }
        private string _MaLoaiMonAn;
        public string MaLoaiMonAn
        {
            get { return _MaLoaiMonAn; }
            set { _MaLoaiMonAn = value; }
        }
        private bool _TinhTrangMonAn;
        public bool TinhTrangMonAn
        {
            get { return _TinhTrangMonAn; }
            set { _TinhTrangMonAn = value; }
        }

        private Byte[] _HinhAnh;
        public Byte[] HinhAnh 
        {
            get { return _HinhAnh; } 
            set { _HinhAnh = value; } 
        }
    }
}
