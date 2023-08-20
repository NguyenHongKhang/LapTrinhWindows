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

namespace WindowsFormsApp1.Views
{
    public partial class AddCategory : Form
    {
        public AddCategory()
        {
            InitializeComponent();
            tbCategoryID.Text = BLL_QLCF.Instance.GetNewCategoryrId();
        }

        private void bCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
        private bool InsertCategory()

        {
            Category kh = new Category(tbCategoryID.Text, tbCategoryName.Text);
            bool result;
            try
            {
                result = BLL_QLCF.Instance.InsertCategory(kh);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (result)
            {
                MessageBox.Show("Thêm loại thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình thêm loại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return true;
        }
        private void bOK_Click(object sender, EventArgs e)
        {
            if (InsertCategory()==true)
                InsertCategory();
            else return;
            this.Dispose();
        }

        private void tbCategoryID_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
