using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.DTO
{
    public class Bill
    {
        public Bill(string TenMonAn, int SoLuong, float Price, float ThanhTien = 0)
        {
            this.TenMonAn = TenMonAn;
            this.SoLuong = SoLuong;
            this.Price = Price;
            this.ThanhTien = ThanhTien;
        }
        public Bill(DataRow row)
        {
            this.TenMonAn = row["TenMonAn"].ToString();
            this.SoLuong = (int)row["SoLuong"];
            this.Price = (float)Convert.ToDouble(row["GiaTien"].ToString());
            this.ThanhTien = (float)Convert.ToDouble(row["ThanhTien"].ToString());
        }
        private string _TenMonAn;

        public string TenMonAn 

        {
            get { return _TenMonAn; }
            set { _TenMonAn = value; }
        }
        private int _SoLuong;
        public int SoLuong
        {
            get { return _SoLuong; }
            set { _SoLuong = value; }
        }
        private float price;
        public float Price 
        {
            get { return price; } 
            set { price = value; } 
        }
        private float _ThanhTien;
        public float ThanhTien 
        {
            get { return _ThanhTien; }
            set { _ThanhTien = value; }
        }



    }
}
