using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.DTO
{
    public class Table
    {
        public Table(string maban, string tenban, string trangthai)
        {
            this.MaBan = maban;

            this.TenBan = tenban;
            this.TrangThai = trangthai;
        }
        public Table(DataRow row)
        {
            this.MaBan = row["MaBan"].ToString();
            this.TenBan = (string)row["TenBan"];
            this.TrangThai = (string)row["TrangThai"];
        }
        private string _MaBan;
        public string MaBan
        {
            get { return _MaBan; }
            set { _MaBan = value; }
        }
        private string _TenBan;
        public string TenBan
        {
            get { return _TenBan; }
            set { _TenBan = value; }
        }
        private string _TrangThai;
        public string TrangThai
        {
            get { return _TrangThai; }
            set { _TrangThai = value; }
        }

    }
}
