using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using WindowsFormsApp1.BLL;
using WindowsFormsApp1.DTO;

namespace WindowsFormsApp1
{
    public partial class Statistics : Form
    {
        public Statistics()
        {
            InitializeComponent();
            LoadCBBItems();
            LoadDtp();
            TDT_ShowStatistics();
            NV_ShowStatistics();
            MA_ShowStatistics();
        }

        #region Methods

        private void LoadDtp()
        {
            TDT_dtpTo.Value = DateTime.Now;
            TDT_dtpFrom.Value = TDT_dtpTo.Value.AddMonths(-1);
            NV_dtpTo.Value = DateTime.Now;
            NV_dtpFrom.Value = TDT_dtpTo.Value.AddMonths(-1);
            MA_dtpTo.Value = DateTime.Now;
            MA_dtpFrom.Value = TDT_dtpTo.Value.AddMonths(-1);

        }

        private void LoadCBBItems()
        {
            CBBItems day = new CBBItems();
            day.Value = 0;
            day.Text = "Ngày";
            TDT_cbStatisticsMode.Items.Add(day);
            CBBItems month = new CBBItems();
            month.Value = 1;
            month.Text = "Tháng";
            TDT_cbStatisticsMode.Items.Add(month);
            CBBItems year = new CBBItems();
            year.Value = 2;
            year.Text = "Năm";
            TDT_cbStatisticsMode.Items.Add(year);
            TDT_cbStatisticsMode.SelectedIndex = 0;
            NV_cbShiftMode.Items.Add(new CBBItems { Text = "All", Value = 0 });
            NV_cbShiftMode.Items.Add(new CBBItems { Text = "Sáng", Value = 1 });
            NV_cbShiftMode.Items.Add(new CBBItems { Text = "Chiều", Value = 2 });
            NV_cbShiftMode.Items.Add(new CBBItems { Text = "Tối", Value = 3 });
            NV_cbShiftMode.SelectedIndex = 0;
            MA_cbCategoriesMode.Items.Add(new CBBItem { Text = "All", Value = "0" });
            foreach(Category i in BLL_QLCF.Instance.GetAllCategories())
            {
                MA_cbCategoriesMode.Items.Add(new CBBItem { Text = i.TenLoaiMonAn, Value = i.MaLoaiMonAn });
            }
            MA_cbCategoriesMode.SelectedIndex = 0;

            CBBItems Cot = new CBBItems();
            Cot.Value = 0;
            Cot.Text = " Biểu đồ cột đơn ";
            TDT_cbChartMode.Items.Add(Cot);
            NV_cbChartMode.Items.Add(Cot);
            MA_cbChartMode.Items.Add(Cot);
            CBBItems CotDung = new CBBItems();
            CotDung.Value = 1;
            CotDung.Text = "Biểu đồ cột ngang ";
            TDT_cbChartMode.Items.Add(CotDung);
            NV_cbChartMode.Items.Add(CotDung);
            MA_cbChartMode.Items.Add(CotDung);
            CBBItems Tron = new CBBItems();
            Tron.Value = 2;
            Tron.Text = "Biểu đồ tròn ";
            TDT_cbChartMode.Items.Add(Tron);
            NV_cbChartMode.Items.Add(Tron);
            MA_cbChartMode.Items.Add(Tron);
            CBBItems Line = new CBBItems();
            Line.Value = 3;
            Line.Text = "Biểu đồ đường";
            TDT_cbChartMode.Items.Add(Line);
            NV_cbChartMode.Items.Add(Line);
            MA_cbChartMode.Items.Add(Line);
            TDT_cbChartMode.SelectedIndex = 0;
            NV_cbChartMode.SelectedIndex = 0;
            MA_cbChartMode.SelectedIndex = 0;
        }

        private void TDT_ShowStatistics()
        {
            TDT_cStatistics.Controls.Clear();
            string Time_min = "";
            string Time_max = "";
            string dinhdangthoigian = "";

            if (TDT_cbStatisticsMode.SelectedIndex == 0)
            {
                Time_min += TDT_dtpFrom.Value.Year + "-" + TDT_dtpFrom.Value.Month + "-" + TDT_dtpFrom.Value.Day + " 00:00:00";
                Time_max += TDT_dtpTo.Value.Year + "-" + TDT_dtpTo.Value.Month + "-" + TDT_dtpTo.Value.Day + " 23:59:59";
                dinhdangthoigian = "dd/MM/yyyy";
            }
            if (TDT_cbStatisticsMode.SelectedIndex == 1)
            {
                Time_min += TDT_dtpFrom.Value.Year + "-" + TDT_dtpFrom.Value.Month + "-01 00:00:00";
                Time_max += TDT_dtpTo.Value.Year + "-" + TDT_dtpTo.Value.Month + "-" + DateTime.DaysInMonth(TDT_dtpTo.Value.Year, TDT_dtpTo.Value.Month).ToString() + " 23:59:59";
                dinhdangthoigian = "MM/yyyy";
            }
            if (TDT_cbStatisticsMode.SelectedIndex == 2)
            {
                Time_min += TDT_dtpFrom.Value.Year + "-01-01 00:00:00";
                Time_max += TDT_dtpTo.Value.Year + "-12-31 23:59:59";
                dinhdangthoigian = "yyyy";
            }

            if (TDT_dtpFrom.Value > TDT_dtpTo.Value)
            {
                MessageBox.Show("Thời gian không hợp lệ!!!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataTable dt = BLL_QLCF.Instance.TDT_GetRevenueTable(dinhdangthoigian, Time_max, Time_min);
            TDT_dgvStatistics.DataSource = dt;
            double tong = 0;
            object temp = dt.Compute("SUM(TongTien)", string.Empty);
            if(typeof(DBNull)!=temp.GetType())
            {
                tong = Convert.ToDouble(temp);
            }  
            TDT_tbTotalRevenue.Text = tong.ToString();

            TDT_cStatistics.DataSource = dt;
            if (TDT_cbChartMode.SelectedIndex != 2)
            {
                Draw_Chart_Column(TDT_cStatistics, TDT_cbChartMode, "ThoiGian", "TongTien", "Thoi Gian", "VND");
            }
            else
            {
                Draw_Chart_Pie(TDT_cStatistics, "ThoiGian", "TongTien");
            }
            
        }
        
        private void NV_ShowStatistics()
        {
            NV_cStatistics.Controls.Clear();
            string Time_max = NV_dtpTo.Value.Year + "-" + NV_dtpTo.Value.Month + "-" + NV_dtpTo.Value.Day + " 23:59:59";
            string Time_min = NV_dtpFrom.Value.Year + "-" + NV_dtpFrom.Value.Month + "-" + NV_dtpFrom.Value.Day + " 00:00:00";
            string Shift = "";
            if(NV_cbShiftMode.SelectedIndex == 1)
            {
                Shift = "Sáng";
            }
            if (NV_cbShiftMode.SelectedIndex == 2)
            {
                Shift = "Chiều";
            }
            if (NV_cbShiftMode.SelectedIndex == 3)
            {
                Shift = "Tối";
            }
            if (TDT_dtpFrom.Value > TDT_dtpTo.Value)
            {
                MessageBox.Show("Thời gian không hợp lệ!!!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataTable dt = BLL_QLCF.Instance.NV_GetRevenueTable(Shift, Time_max, Time_min);
            NV_dgvStatistics.DataSource = dt;
            double tong = 0;
            object temp = dt.Compute("SUM(DoanhThu)", string.Empty);
            if (typeof(DBNull) != temp.GetType())
            {
                tong = Convert.ToDouble(temp);
            }
            NV_tbTotalRevenue.Text = tong.ToString();

            NV_cStatistics.DataSource = dt;
            if (NV_cbChartMode.SelectedIndex != 2)
            {
                Draw_Chart_Column(NV_cStatistics, NV_cbChartMode, "TenNhanVien", "DoanhThu", "Nhân Viên", "VND");
            }
            else
            {
                Draw_Chart_Pie(NV_cStatistics, "TenNhanVien", "DoanhThu");
            }

        }

        private void MA_ShowStatistics()
        {
            MA_cStatistics.Controls.Clear();
            string Time_max = MA_dtpTo.Value.Year + "-" + MA_dtpTo.Value.Month + "-" + MA_dtpTo.Value.Day + " 23:59:59";
            string Time_min = MA_dtpFrom.Value.Year + "-" + MA_dtpFrom.Value.Month + "-" + MA_dtpFrom.Value.Day + " 00:00:00";
            string Category = "";
            if(!((CBBItem)MA_cbCategoriesMode.SelectedItem).Text.Equals("All"))
            {
                Category = ((CBBItem)MA_cbCategoriesMode.SelectedItem).Value;
            }
            if (MA_dtpFrom.Value > MA_dtpTo.Value)
            {
                MessageBox.Show("Thời gian không hợp lệ!!!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataTable dt = BLL_QLCF.Instance.MA_GetRevenueTable(Category, Time_max, Time_min);
            MA_dgvStatistics.DataSource = dt;
            double tong = 0;
            object temp = dt.Compute("SUM(DoanhThu)", string.Empty);
            if (typeof(DBNull) != temp.GetType())
            {
                tong = Convert.ToDouble(temp);
            }
            MA_tbTotalRevenue.Text = tong.ToString();

            MA_cStatistics.DataSource = dt;
            if (MA_cbChartMode.SelectedIndex != 2)
            {
                Draw_Chart_Column(MA_cStatistics, MA_cbChartMode, "TenMonAn", "DoanhThu", "Món Ăn", "VND");
            }
            else
            {
                Draw_Chart_Pie(MA_cStatistics, "TenMonAn", "DoanhThu");
            }

        }

        private void Draw_Chart_Column(Chart Ch , ComboBox Cb, string X, string Y, string XValue, string YValue)
        {
            Ch.Series.Clear();
            Ch.ChartAreas.Clear();
            Ch.ChartAreas.Add("ChartArea1");
            Ch.ChartAreas["ChartArea1"].AxisX.Title = XValue;
            Ch.ChartAreas["ChartArea1"].AxisY.Title = YValue;
            Ch.Series.Add("Series1");
            if(Cb.SelectedIndex==0)
            {
                Ch.Series["Series1"].ChartType = SeriesChartType.Column;
            }
            if (Cb.SelectedIndex == 1)
            {
                Ch.Series["Series1"].ChartType = SeriesChartType.Bar;
            }
            if (Cb.SelectedIndex == 3)
            {
                Ch.Series["Series1"].ChartType = SeriesChartType.Line;
            }
            Ch.Series["Series1"].IsXValueIndexed = false;
            Ch.Series["Series1"].XValueMember = X;
            Ch.Series["Series1"].YValueMembers = Y;
            Ch.Series["Series1"].Name = "Tổng tiền";
        }

        private void Draw_Chart_Pie(Chart Ch, string X, string Y)
        {
            Ch.Series.Clear();
            Ch.ChartAreas.Clear();
            Ch.ChartAreas.Add("ChartArea1");
            Ch.Series.Add("Series1");
            Ch.Series["Series1"].ChartType = SeriesChartType.Pie;
            Ch.Series["Series1"].IsXValueIndexed = true;
            Ch.Series["Series1"].XValueMember = X;
            Ch.Series["Series1"].YValueMembers = Y;
            Ch.Series["Series1"].Name = "Value";
        }

        #endregion

        #region Events
        private void TDT_btShow_Click(object sender, EventArgs e)
        {
            TDT_ShowStatistics();
        }
        #endregion

        private void TDT_cbStatisticsMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            string dinhdang = "dd/MM/yyyy";
            if (TDT_cbStatisticsMode.SelectedIndex == 1) dinhdang = "MM/yyyy";
            if (TDT_cbStatisticsMode.SelectedIndex == 2) dinhdang = "yyyy";
            TDT_dtpFrom.Format = DateTimePickerFormat.Custom;
            TDT_dtpFrom.CustomFormat = dinhdang;
            TDT_dtpTo.Format = DateTimePickerFormat.Custom;
            TDT_dtpTo.CustomFormat = dinhdang;
        }

        private void NV_btShow_Click(object sender, EventArgs e)
        {
            NV_ShowStatistics();
        }

        private void MA_btShow_Click(object sender, EventArgs e)
        {
            MA_ShowStatistics();
        }

        private void TDT_cStatistics_Click(object sender, EventArgs e)
        {


        }
    }
}
