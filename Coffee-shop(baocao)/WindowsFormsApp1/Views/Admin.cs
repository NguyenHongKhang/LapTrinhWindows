using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.BLL;
using WindowsFormsApp1.DAL;
using WindowsFormsApp1.DAO;
using WindowsFormsApp1.DTO;
using WindowsFormsApp1.Views;

namespace WindowsFormsApp1
{
    public partial class Admin : Form
    {
        BindingSource DgvFoodBinding;
        BindingSource DgvCategoryBinding;
        BindingSource DgvTableBinding;
        BindingSource DgvStaffBinding;
        BindingSource DgvAccountBinding;

        public Admin()
        {
            InitializeComponent();
            DgvFoodBinding = new BindingSource();
            dgvFood.DataSource = DgvFoodBinding;
            DgvCategoryBinding = new BindingSource();
            dgvCategory.DataSource = DgvCategoryBinding;
            DgvTableBinding = new BindingSource();
            dgvTable.DataSource = DgvTableBinding;
            DgvStaffBinding = new BindingSource();
            dgvNhanVien.DataSource = DgvStaffBinding;
            DgvAccountBinding = new BindingSource();
            dgvTaiKhoan.DataSource = DgvAccountBinding;
            load();
        }

        #region methods
        private void load()
        {
            LoadDateTimePicker();
            LoadListFood();
            LoadListCategory();
            LoadListTable();
            LoadListStaff();
            LoadListAccount();
            SetShop();
        }
        private void LoadListFood()
        {
            SetCBBItem();
            GetCategoryToComboBox();
            AddFoodBinding();
        }
        private void LoadListCategory()
        {
            DgvCategoryBinding.DataSource = BLL_QLCF.Instance.GetAllCategories();
            SetCBBItemCategory();
            AddCategoryBinding();
        }
        private void LoadListTable()
        {
            DgvTableBinding.DataSource = BLL_QLCF.Instance.GetAllTable();
            AddTableBinding();
        }
        private void LoadListStaff()
        {
            DgvStaffBinding.DataSource = BLL_QLCF.Instance.GetAllStaff();
            dgvNhanVien.Columns[7].Visible = false;
            SetCBBItemStaff();
            AddStaffBinding();
        }
        private void LoadListAccount()
        {
            DgvAccountBinding.DataSource = BLL_QLCF.Instance.GetAllAccount();
            SetCBBItemAccount();
            GetStaffIDToComboBox();
            GetAccountTypeToComboBox();
            AddAccountBinding();
        }
        private void SetShop()
        {
            AddressShop shop = BLL_QLCF.Instance.GetShop();
            tbNameShop.Text = shop.TenShop;
            tbAddress.Text = shop.DiaChiShop;
            tbownShop.Text = shop.TenChuShop;
            tbPhoneShop.Text = shop.SoDienThoai;
        }
        private void AddFoodBinding()
        {
            tbFoodID.DataBindings.Add(new Binding("Text", dgvFood.DataSource, "MaMonAn", true, DataSourceUpdateMode.Never));
            tbFoodsName.DataBindings.Add(new Binding("Text", dgvFood.DataSource, "TenMonAn", true, DataSourceUpdateMode.Never));
            nudFoodPrice.DataBindings.Add(new Binding("Value", dgvFood.DataSource, "GiaTien", true, DataSourceUpdateMode.Never));
        }
        private void AddCategoryBinding()
        {
            tbCategoryID.DataBindings.Add(new Binding("Text", dgvCategory.DataSource, "MaLoaiMonAn", true, DataSourceUpdateMode.Never));
            tbCategoryName.DataBindings.Add(new Binding("Text", dgvCategory.DataSource, "TenLoaiMonAn", true, DataSourceUpdateMode.Never));
        }
        private void AddTableBinding()
        {
            tbMaBan.DataBindings.Add(new Binding("Text", dgvTable.DataSource, "MaBan", true, DataSourceUpdateMode.Never));
            tbTenBan.DataBindings.Add(new Binding("Text", dgvTable.DataSource, "TenBan", true, DataSourceUpdateMode.Never));
            tbTrangThai.DataBindings.Add(new Binding("Text", dgvTable.DataSource, "TrangThai", true, DataSourceUpdateMode.Never));
        }
        private void AddStaffBinding()
        {
            tbMaNhanVien.DataBindings.Add(new Binding("Text", dgvNhanVien.DataSource, "MaNhanVien", true, DataSourceUpdateMode.Never));
            tbTenNhanVien.DataBindings.Add(new Binding("Text", dgvNhanVien.DataSource, "TenNhanVien", true, DataSourceUpdateMode.Never));
            tbSDT.DataBindings.Add(new Binding("Text", dgvNhanVien.DataSource, "SoDienThoai", true, DataSourceUpdateMode.Never));
            tbDC.DataBindings.Add(new Binding("Text", dgvNhanVien.DataSource, "DiaChi", true, DataSourceUpdateMode.Never));
            tbCMND.DataBindings.Add(new Binding("Text", dgvNhanVien.DataSource, "CMND", true, DataSourceUpdateMode.Never));
            tbQueQuan.DataBindings.Add(new Binding("Text", dgvNhanVien.DataSource, "QueQuan", true, DataSourceUpdateMode.Never));
            tbCaLamViec.DataBindings.Add(new Binding("Text", dgvNhanVien.DataSource, "CaLamViec", true, DataSourceUpdateMode.Never));
        }
        private void AddAccountBinding()
        {
            tbTenTaiKhoan.DataBindings.Add(new Binding("Text", dgvTaiKhoan.DataSource, "TenTaiKhoan", true, DataSourceUpdateMode.Never));
            tbMatKhau.DataBindings.Add(new Binding("Text", dgvTaiKhoan.DataSource, "MatKhau", true, DataSourceUpdateMode.Never));
            tbTenHienThi.DataBindings.Add(new Binding("Text", dgvTaiKhoan.DataSource, "TenHienThi", true, DataSourceUpdateMode.Never));
        }
        private void SetCBBItemCategory()
        {
            cbSortCategory.Items.Add(new CBBItem { Text = "Tăng theo tên loại món ăn", Value = "1" });
            cbSortCategory.Items.Add(new CBBItem { Text = "Giảm theo tên loại món ăn", Value = "2" });

            cbSortCategory.SelectedItem = cbSortCategory.Items[0];
        }
        private void SetCBBItemStaff()
        {
            cbSortStaff.Items.Add(new CBBItem { Text = "Tăng theo tên nhân viên", Value = "1" });
            cbSortStaff.Items.Add(new CBBItem { Text = "Giảm theo tên nhân viên", Value = "2" });

            cbSortStaff.SelectedItem = cbSortStaff.Items[0];
        }
        private void SetCBBItem()
        {
            cbcategory.Items.Clear();
            cbcategory.Items.Add(new CBBItem { Text = "Tất cả", Value = "0" });
            cbcategory.Items.AddRange(BLL_QLCF.Instance.GetCategoryCBBItem().ToArray());

            cbsort.Items.Add(new CBBItem { Text = "Tăng theo tên món ăn", Value = "1" });
            cbsort.Items.Add(new CBBItem { Text = "Giảm theo tên món ăn", Value = "2" });

            cbsort.SelectedItem = cbsort.Items[0];
            cbcategory.SelectedItem = cbcategory.Items[0];
        }
        private void SetCBBItemAccount()
        {
            cbLoaiTK.Items.Clear();
            cbLoaiTK.Items.Add(new CBBItems { Text = "Tất cả", Value = 2 });
            cbLoaiTK.Items.AddRange(BLL_QLCF.Instance.GetAccountTypeCBBItem().ToArray());

            cbSortTK.Items.Add(new CBBItem { Text = "Tăng theo tên tài khoản", Value = "1" });
            cbSortTK.Items.Add(new CBBItem { Text = "Giảm theo tên tài khoản", Value = "2" });

            cbSortTK.SelectedItem = cbSortTK.Items[0];
            cbLoaiTK.SelectedItem = cbLoaiTK.Items[0];
        }
        private void ShowCategory()
        {
            string ValueSort = ((CBBItem)cbSortCategory.SelectedItem).Value;
            string KeyWord = tbFindCategory.Text;
            DgvCategoryBinding.DataSource = BLL_QLCF.Instance.GetListCategory(ValueSort, KeyWord);
        }
        private void ShowTable()
        {
            string KeyWord = tbFindTable.Text;
            DgvTableBinding.DataSource = BLL_QLCF.Instance.FindTable(KeyWord);
        }
        private void ShowFood()
        {
            string ValueCategory = ((CBBItem)cbcategory.SelectedItem).Value;
            string ValueSort = ((CBBItem)cbsort.SelectedItem).Value;
            string KeyWord = tbfind.Text;
            DgvFoodBinding.DataSource = BLL_QLCF.Instance.GetListFood(ValueCategory, ValueSort, KeyWord);
            dgvFood.Columns[5].Visible = false;
        }
        private void ShowStaff()
        {
            string ValueSort = ((CBBItem)cbSortStaff.SelectedItem).Value;
            string KeyWord = tbFindStaff.Text;
            DgvStaffBinding.DataSource = BLL_QLCF.Instance.GetListStaff(ValueSort, KeyWord);
            dgvNhanVien.Columns[7].Visible = false;
        }
        private void ShowAccount()
        {
            int ValueCategory = ((CBBItems)cbLoaiTK.SelectedItem).Value;
            string ValueSort = ((CBBItem)cbSortTK.SelectedItem).Value;
            string KeyWord = tbFindTK.Text;
            DgvAccountBinding.DataSource = BLL_QLCF.Instance.GetListAccount(ValueCategory, ValueSort, KeyWord);
        }
        private void GetCategoryToComboBox()
        {
            cbFoodCategory.DataSource = BLL_QLCF.Instance.GetAllCategories();
            cbFoodCategory.DisplayMember = "TenLoaiMonAn";
        }
        private void GetStaffIDToComboBox()
        {
            cbMaNhanVien.DataSource = BLL_QLCF.Instance.GetAllStaff();
            cbMaNhanVien.DisplayMember = "MaNhanVien";
        }
        private void GetAccountTypeToComboBox()
        {
            cbLoaiTaiKhoan.DataSource = BLL_QLCF.Instance.GetAccountTypeCBBItem();
        }
        private void UpdateFood()
        {
            byte[] HinhAnh;
            ImageConverter converter = new ImageConverter();
            HinhAnh = (byte[])converter.ConvertTo(pbfood.Image, typeof(byte[]));
            DoAn da = new DoAn(tbFoodsName.Text, tbFoodID.Text, (float)nudFoodPrice.Value, ((Category)cbFoodCategory.SelectedItem).MaLoaiMonAn, btrue.Checked, HinhAnh);
            bool result;
            try
            {
                result = BLL_QLCF.Instance.UpdateFood(da);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (result)
            {
                MessageBox.Show("Cập nhật thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình cập nhật", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return;
        }
        private void UpdateCategory()
        {
            Category da = new Category(tbCategoryID.Text, tbCategoryName.Text);
            bool result;
            try
            {
                result = BLL_QLCF.Instance.UpdateCategory(da);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (result)
            {
                MessageBox.Show("Cập nhật thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình cập nhật", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return;
        }
        private void UpdateTable()
        {
            Table da = new Table(tbMaBan.Text, tbTenBan.Text, tbTrangThai.Text);
            bool result;
            try
            {
                result = BLL_QLCF.Instance.UpdateTable(da);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (result)
            {
                MessageBox.Show("Cập nhật thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình cập nhật", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return;
        }
        private void UpdateAccount()
        {
            Account da = new Account(tbTenTaiKhoan.Text,tbMatKhau.Text,tbTenHienThi.Text,((CBBItems)cbLoaiTaiKhoan.SelectedItem).Value,cbMaNhanVien.Text);
            bool result;
            try
            {
                result = BLL_QLCF.Instance.UpdateAccount(da);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (result)
            {
                MessageBox.Show("Cập nhật thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình cập nhật", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return;
        }
        private void UpdateStaff()
        {
            byte[] HinhAnh;
            ImageConverter converter = new ImageConverter();
            HinhAnh = (byte[])converter.ConvertTo(pbNhanVien.Image, typeof(byte[]));
            NhanVien da = new NhanVien(tbTenNhanVien.Text, tbMaNhanVien.Text, tbSDT.Text, tbDC.Text, tbCMND.Text, tbQueQuan.Text, tbCaLamViec.Text, HinhAnh);
            bool result;
            try
            {
                result = BLL_QLCF.Instance.UpdateStaff(da);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (result)
            {
                MessageBox.Show("Cập nhật thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình cập nhật", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return;
        }
        private void DeleteFood()
        {
            List<string> ListDelete = new List<string>();
            foreach(DataGridViewRow i in dgvFood.SelectedRows)
            {
                ListDelete.Add(i.Cells["MaMonAn"].Value.ToString());
            }
            if (BLL_QLCF.Instance.DeleteFood(ListDelete))
            {
                MessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình xóa", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void DeleteCategory()
        {
            List<string> ListDelete = new List<string>();
            foreach (DataGridViewRow i in dgvCategory.SelectedRows)
            {
                ListDelete.Add(i.Cells["MaLoaiMonAn"].Value.ToString());
            }
            if (BLL_QLCF.Instance.DeleteCategory(ListDelete))
            {
                MessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình xóa", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void DeleteTable()
        {
            List<string> ListDelete = new List<string>();
            foreach (DataGridViewRow i in dgvTable.SelectedRows)
            {
                ListDelete.Add(i.Cells["MaBan"].Value.ToString());
            }
            if (BLL_QLCF.Instance.DeleteTable(ListDelete))
            {
                MessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình xóa", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void DeleteStaff()
        {
            List<string> ListDelete = new List<string>();
            foreach (DataGridViewRow i in dgvNhanVien.SelectedRows)
            {
                ListDelete.Add(i.Cells["MaNhanVien"].Value.ToString());
            }
            if (BLL_QLCF.Instance.DeleteStaff(ListDelete))
            {
                MessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình xóa", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void DeleteAccount()
        {
            List<string> ListDelete = new List<string>();
            foreach (DataGridViewRow i in dgvTaiKhoan.SelectedRows)
            {
                ListDelete.Add(i.Cells["TenTaiKhoan"].Value.ToString());
            }
            if (BLL_QLCF.Instance.DeleteAccount(ListDelete))
            {
                MessageBox.Show("Xóa thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình xóa", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        void LoadDateTimePicker()
        {
            DateTime today = DateTime.Now;
            dateTimePicker1.Value = new DateTime(today.Year, today.Month, 1);
            dateTimePicker2.Value = dateTimePicker1.Value.AddMonths(1).AddDays(-1);
        }
                     
        #endregion

        #region events
        private void btViewBills_Click(object sender, EventArgs e)
        {
            dgvBills.DataSource = BLL_QLCF.Instance.LoadListBillByDate(dateTimePicker1.Value, dateTimePicker2.Value);
        }
        private void tbFoodID_TextChanged(object sender, EventArgs e)
        {
            string MaMonAn = tbFoodID.Text;
            DoAn food = BLL_QLCF.Instance.GetFoodByFoodId(MaMonAn);
            if (food != null)
            {
                if (food.TinhTrangMonAn == true)
                {
                    btrue.Checked = true;
                }
                else
                {
                    bFalse.Checked = true;
                }
                
                foreach(Category i in cbFoodCategory.Items)
                {
                    if (i.MaLoaiMonAn.Equals(food.MaLoaiMonAn))
                    {
                        cbFoodCategory.SelectedItem = i;
                        break;
                    }
                }

                if (food.HinhAnh != null)
                {
                    MemoryStream ms = new MemoryStream(food.HinhAnh);
                    pbfood.Image = Image.FromStream(ms);
                }
                else
                {
                    pbfood.Image = null;
                }
            }      
        }
        private void tbMaNhanVien_TextChanged(object sender, EventArgs e)
        {
            string MaNhanVien = tbMaNhanVien.Text;
            NhanVien food = BLL_QLCF.Instance.GetStaffByStaffId(MaNhanVien);
            if (food != null)
            {
                if (food.HinhAnh != null)
                {
                    MemoryStream ms = new MemoryStream(food.HinhAnh);
                    pbNhanVien.Image = Image.FromStream(ms);
                }
                else
                {
                    pbNhanVien.Image = null;
                }
            }
        }
        private void tbTenTaiKhoan_TextChanged(object sender, EventArgs e)
        {
            string TenTaiKhoan = tbTenTaiKhoan.Text;
            Account food = BLL_QLCF.Instance.GetAccountByAccountName(TenTaiKhoan);
            if (food != null)
            {
                foreach (CBBItems i in cbLoaiTaiKhoan.Items)
                {
                    if (i.Value.Equals(food.LoaiTaiKhoan))
                    {
                        cbLoaiTaiKhoan.SelectedItem = i;
                        break;
                    }
                }
                foreach (NhanVien i in cbMaNhanVien.Items)
                {
                    if (i.MaNhanVien.Equals(food.MaNhanVien))
                    {
                        cbMaNhanVien.SelectedItem = i;
                        break;
                    }
                }
            }
        }
        private void buttonAddFood_Click(object sender, EventArgs e)
        {
            AddFood f1 = new AddFood();
            f1.ShowDialog();
            ShowFood();
        }
        private void buttonFixFood_Click(object sender, EventArgs e)
        {
            UpdateFood();
            ShowFood();
        }
        private void buttonDeleteFood_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có thật sự muốn xóa " + dgvFood.SelectedRows.Count.ToString() + " món ăn?", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                DeleteFood();
                ShowFood();
            }
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
                        pbfood.Image = new Bitmap(openFileDialog1.FileName);
                        pbfood.SizeMode = PictureBoxSizeMode.StretchImage;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void pbNhanVien_Click(object sender, EventArgs e)
        {
            openFileDialog2.InitialDirectory = "C://Desktop";
            openFileDialog2.Title = "Select image to be upload.";
            openFileDialog2.Filter = "Image Only(*.jpg; *.jpeg; *.gif; *.bmp; *.png)|*.jpg; *.jpeg; *.gif; *.bmp; *.png";
            openFileDialog2.FilterIndex = 1;
            try
            {
                if (openFileDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (openFileDialog2.CheckFileExists)
                    {
                        pbNhanVien.Image = new Bitmap(openFileDialog2.FileName);
                        pbNhanVien.SizeMode = PictureBoxSizeMode.StretchImage;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void cbcategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            ShowFood();
        }
        private void btsort_Click(object sender, EventArgs e)
        {
            ShowFood();
        }
        private void btfind_Click(object sender, EventArgs e)
        {
            ShowFood();
        }
        private void btSortCategory_Click(object sender, EventArgs e)
        {
            ShowCategory();
        }
        private void btFindCategory_Click(object sender, EventArgs e)
        {
            ShowCategory();
        }
        private void btAddCategory_Click(object sender, EventArgs e)
        {
            AddCategory f2 = new AddCategory();
            f2.ShowDialog();
            SetCBBItem();
            GetCategoryToComboBox();
            ShowCategory();
        }
        private void btFixCategory_Click(object sender, EventArgs e)
        {
            UpdateCategory();
            SetCBBItem();
            GetCategoryToComboBox();
            ShowCategory();
        }
        private void btDeleteCategory_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có thật sự muốn xóa " + dgvCategory.SelectedRows.Count.ToString() + " loại món ăn?", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                DeleteCategory();
                ShowFood();
                SetCBBItem();
                GetCategoryToComboBox();
                ShowCategory();
            }
        }
        private void btAddTable_Click(object sender, EventArgs e)
        {
            AddTable f2 = new AddTable();
            f2.ShowDialog();
            ShowTable();
        }
        private void btFindTable_Click(object sender, EventArgs e)
        {
            ShowTable();
        }
        private void btFixTable_Click(object sender, EventArgs e)
        {
            UpdateTable();
            ShowTable();
        }
        private void btDelTable_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có thật sự muốn xóa " + dgvTable.SelectedRows.Count.ToString() + " bàn ăn?", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                DeleteTable();
                ShowTable();
            }
        }
        private void btAddStaff_Click(object sender, EventArgs e)
        {
            AddStaff f2 = new AddStaff();
            f2.ShowDialog();
            GetStaffIDToComboBox();
            ShowStaff();
        }
        private void btFixStaff_Click(object sender, EventArgs e)
        {
            UpdateStaff();
            ShowStaff();
        }
        private void btDelStaff_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có thật sự muốn xóa " + dgvNhanVien.SelectedRows.Count.ToString() + " nhân viên?", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                DeleteStaff();
                ShowStaff();
            }
        }
        private void btSortStaff_Click(object sender, EventArgs e)
        {
            ShowStaff();
        }
        private void btFindStaff_Click(object sender, EventArgs e)
        {
            ShowStaff();
        }
        private void btAddAccount_Click(object sender, EventArgs e)
        {
            AddAccount f2 = new AddAccount();
            f2.ShowDialog();
            ShowAccount();
        }
        private void btFixAccount_Click(object sender, EventArgs e)
        {
            UpdateAccount();
            ShowAccount();
        }
        private void btDelAccount_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có thật sự muốn xóa " + dgvTaiKhoan.SelectedRows.Count.ToString() + " tài khoản?", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                DeleteAccount();
                ShowAccount();
            }
        }
        private void btSortTK_Click(object sender, EventArgs e)
        {
            ShowAccount();
        }
        private void btFindTK_Click(object sender, EventArgs e)
        {
            ShowAccount();
        }
        private void cbLoaiTK_SelectedIndexChanged(object sender, EventArgs e)
        {
            ShowAccount();
        }
        private void btShopUpdate_Click(object sender, EventArgs e)
        {
            AddressShop shop = new AddressShop(tbNameShop.Text, tbownShop.Text, tbAddress.Text, tbPhoneShop.Text);
            string TenShopNew = tbNameShop.Text;
            string DiaChiNew = tbAddress.Text;
            string ChuNew = tbownShop.Text;
            string SoDienThoaiNew = tbPhoneShop.Text;
            bool result;
            try
            {
                result = BLL_QLCF.Instance.UpdateShop(shop, TenShopNew, DiaChiNew, ChuNew, SoDienThoaiNew);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (result == false)
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình cập nhật", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Cập nhật thành công!!!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return;
        }

        #endregion

        private void dgvBills_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


        }
    }
}
