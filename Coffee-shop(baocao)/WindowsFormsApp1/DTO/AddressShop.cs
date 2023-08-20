using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.DTO
{
    class AddressShop
    {
        public AddressShop(string TenShop, string TenChuShop, string DiaChiShop, string SoDienThoai)
        {
            this.TenShop = TenShop;

            this.TenChuShop = TenChuShop;
            this.DiaChiShop = DiaChiShop;
            this.SoDienThoai = SoDienThoai;
        }
        public AddressShop(DataRow row)
        {
            this.TenShop = row["TenShop"].ToString();
            this.TenChuShop = row["TenChuShop"].ToString();
            this.DiaChiShop = row["DiaChiShop"].ToString();
            this.SoDienThoai = row["SoDienThoai"].ToString();
        }
        private string _TenShop;

        public string TenShop
        {
            get { return _TenShop; }
            set { _TenShop = value; }
        }

        private string _TenChuShop;

        public string TenChuShop
        {
            get { return _TenChuShop; }
            set { _TenChuShop = value; }
        }
        private string _DiaChiShop;

        public string DiaChiShop
        {
            get { return _DiaChiShop; }
            set { _DiaChiShop = value; }
        }
        private string _SoDienThoai;
        public string SoDienThoai
        {
            get { return _SoDienThoai; }
            set { _SoDienThoai = value; }
        }
    }
}
