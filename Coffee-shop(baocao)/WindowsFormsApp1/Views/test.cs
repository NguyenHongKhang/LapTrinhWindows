using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;
using System.Windows.Forms;

namespace WindowsFormsApp1.Views
{
    public partial class test : Form
    {
        public test()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Khởi tạo đối tượng Document
            Document doc = new Document();

            // Đường dẫn tới thư mục chứa tệp PDF
            string outputPath = @"E:\BT C#\Coffee-shop(baocao)\Hoá đơn\PDF\";

            // Tên tệp PDF bạn muốn tạo
            string fileName = "my_pdf_file.pdf";

            // Đường dẫn hoàn chỉnh của tệp PDF
            string completePath = Path.Combine(outputPath, fileName);

            // Tạo tệp mới
            using (FileStream fs = new FileStream(completePath, FileMode.Create))
            {
                // Tạo tài liệu PDF
                PdfWriter writer = PdfWriter.GetInstance(doc, fs);

                // Mở tài liệu để bắt đầu viết
                doc.Open();

                // Thêm nội dung vào tài liệu
                doc.Add(new Paragraph("Xin chào"));

                // Đóng tài liệu
                doc.Close();
            }

            MessageBox.Show("Tạo tệp PDF thành công!");
        }

    }
}
