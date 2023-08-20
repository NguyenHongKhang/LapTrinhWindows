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
using WindowsFormsApp1.DAO;
using WindowsFormsApp1.DTO;
using WindowsFormsApp1.Properties;

namespace WindowsFormsApp1
{
    public partial class AddFood : Form
    {
        public AddFood()
        {
            InitializeComponent();
            LoadForm();
        }
        private void GetCategoryToComboBox()
        {
            cbFoodCategory.DataSource = BLL_QLCF.Instance.GetAllCategories();
            cbFoodCategory.DisplayMember = "TenLoaiMonAn";
        }
        private void LoadForm()
        {
            GetCategoryToComboBox();
            btrue.Checked = true;
            tbFoodID.Text = BLL_QLCF.Instance.GetNewFoodId();
            tbFoodID.Enabled = false;
            pbfood.BackgroundImage = Resources.no_photo;
            pbfood.BackgroundImageLayout = ImageLayout.Stretch;
        }

        private bool InsertFood()
        {
            byte[] HinhAnh;
            ImageConverter converter = new ImageConverter();
            HinhAnh = (byte[])converter.ConvertTo(pbfood.BackgroundImage, typeof(byte[]));
            DoAn da = new DoAn(tbFoodsName.Text, tbFoodID.Text, (float)nudFoodPrice.Value, ((Category)cbFoodCategory.SelectedItem).MaLoaiMonAn, btrue.Checked, HinhAnh);
            bool result;
            try
            {
                result = BLL_QLCF.Instance.InsertFood(da);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (result)
            {
                MessageBox.Show("Thêm món thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình thêm món", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return true;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
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
            if (InsertFood() == true)
                this.Dispose();
            else return;
        }

        private void bCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}
