using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.DTO;

namespace WindowsFormsApp1.Views
{
    public partial class AccountView : Form
    {
        #region Assembly Attribute Accessors

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion
        private NhanVien _nhanvien;

        public NhanVien nhanvien
        {
            get { return _nhanvien; }
            set { _nhanvien = value; SetAccount(nhanvien); }
        }
        public AccountView(NhanVien nhanvien)
        {
            InitializeComponent();
            this.nhanvien = nhanvien;
        }
        void SetAccount(NhanVien nhanvien)
        {
            tbName.Text = nhanvien.TenNhanVien;
            tbSeries.Text = nhanvien.MaNhanVien;
            tbPhone.Text = nhanvien.SoDienThoai;
            tbCMND.Text = nhanvien.CMND;
            tbAddress.Text = nhanvien.DiaChi;
            tbCountry.Text = nhanvien.QueQuan;
            tbCalam.Text = nhanvien.CaLamViec;
            MemoryStream ms = new MemoryStream(nhanvien.HinhAnh);
            pictureBox1.Image = Image.FromStream(ms);
        }

        private void bOK_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }
}
