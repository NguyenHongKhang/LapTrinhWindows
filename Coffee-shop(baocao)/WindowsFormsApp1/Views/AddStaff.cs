using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.BLL;
using WindowsFormsApp1.DTO;
using WindowsFormsApp1.Properties;

namespace WindowsFormsApp1.Views
{
    public partial class AddStaff : Form
    {
        public AddStaff()
        {
            InitializeComponent();
            LoadForm();
        }
        #region methods
        private bool InsertStaff()
        {
            byte[] HinhAnh;
            ImageConverter converter = new ImageConverter();
            HinhAnh = (byte[])converter.ConvertTo(pbfood.BackgroundImage, typeof(byte[]));
            NhanVien da = new NhanVien(tbTenNhanVien.Text, tbMaNhanVien.Text, tbSDT.Text, tbDC.Text, tbCMND.Text, tbQueQuan.Text, tbCaLamViec.Text, HinhAnh);
            bool result;
            try
            {
                result = BLL_QLCF.Instance.InsertStaff(da);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (result)
            {
                MessageBox.Show("Thêm nhân viên thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình thêm nhân viên", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return true;
        }
        private void LoadForm()
        {
            tbMaNhanVien.Text = BLL_QLCF.Instance.GetNewStaffId();
            pbfood.BackgroundImage = Resources.no_photo;
            pbfood.BackgroundImageLayout = ImageLayout.Stretch;
        }
        #endregion
        #region events
        private void pbfood_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "C://Desktop";

            openFileDialog1.Title = "Select image to be upload.";
            openFileDialog1.Filter = "Image Only(*.jpg; *.jpeg; *.gif; *.bmp; *.png)|*.jpg; *.jpeg; *.gif; *.bmp; *.png";
            openFileDialog1.FilterIndex = 1;
            try
            {
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (openFileDialog1.CheckFileExists)
                    {
                        pbfood.BackgroundImage = new Bitmap(openFileDialog1.FileName);
                        pbfood.SizeMode = PictureBoxSizeMode.StretchImage;
                        label2.Hide();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void bOK_Click(object sender, EventArgs e)
        {
            if (InsertStaff() == true)
                this.Dispose();
            else return;
        }

        private void bCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
        #endregion
    }
}
