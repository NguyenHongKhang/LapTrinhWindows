using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.DTO
{
    public class Category
    {
        public Category(string MaLoaiMonAn, string TenLoaiMonAn)
        {
            this.MaLoaiMonAn = MaLoaiMonAn;

            this.TenLoaiMonAn = TenLoaiMonAn;
        }
        public Category(DataRow row)
        {
            this.MaLoaiMonAn = row["MaLoaiMonAn"].ToString();
            this.TenLoaiMonAn = row["TenLoaiMonAn"].ToString();
        }
        private string _MaLoaiMonAn;

        public string MaLoaiMonAn 
        {
            get { return _MaLoaiMonAn; }
            set { _MaLoaiMonAn = value; } 
        }
        private string _TenLoaiMonAn;
        public string TenLoaiMonAn 
        {
            get { return _TenLoaiMonAn; }
            set { _TenLoaiMonAn = value; }
        }
    }
}
