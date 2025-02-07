using QRCoder;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Print_Label
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public int SatirSayisi = 0;
        public int sayac = 0;
        public string Firma = "TEKKAN";
        string formattedDateForQR = DateTime.Now.ToString("ddMMyyyy"); // YYYYMMDD formatında, noktasız
        string formattedTimeForQR = DateTime.Now.ToString("HHmm");   // HHMMSS formatında, noktasız
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs ffff)
        {
            try
            {
                //ÇİZİM BAŞLANGICI
                Font myFont = new Font("Calibri", 7); //font oluşturduk
                SolidBrush sbrush = new SolidBrush(Color.Black);//fırça oluşturduk
                Pen myPen = new Pen(Color.Black); //kalem oluşturduk

                int x = 18; //x koordinatının yerini belirledik
                int y = 92; //y koordinatının yerini belirledik
                myFont = new Font("Calibri", 6); //fontu 6 yaptık.
                QRCodeGenerator qrGenerator = new QRCodeGenerator();


                int combo = Convert.ToInt32(comboBox1.Text);
                for (sayac = 0; sayac < combo; sayac += 5) //sayac ı global değişken yaptık.
                {
                    StringFormat yaziFormati = new StringFormat(); //yeni bir yazı formatı tanımladık.
                    yaziFormati.FormatFlags = StringFormatFlags.DirectionVertical; // yazı özelliklerini dikey yönde ayarladık
                                                                                   // x ekseni yatayda hizhalamaya yarar
                                                                                   // y ekseni dikeyde hizalamaya yarar

                    for (int i = 0; i < 5; i++)
                    {

                        if (i >= dataGridView1.Rows.Count)
                            break;

                        //string etiketBilgi = dataGridView1[1, i].Value.ToString(); // Stok kodu
                        string etiketBilgi = $"{dataGridView1[1, i].Value.ToString() + "-M"}" +
                            "XX" + //Revizyon
                            "016477" + //arçelik kodu
                            "0900100" + //ülke kodu
                            $"{formattedDateForQR}{formattedTimeForQR}" +
                            "0000000000" +
                            $"{dataGridView1[5, i].Value.ToString()}" +
                            $"{dataGridView1[2, i].Value.ToString()}";

                        QRCodeData qrCodeData = qrGenerator.CreateQrCode(etiketBilgi, QRCodeGenerator.ECCLevel.Q);
                        QRCode qrCode = new QRCode(qrCodeData);
                        Bitmap qrCodeImage = qrCode.GetGraphic(5);


                        ffff.Graphics.DrawImage(qrCodeImage, x + (i * 68) - 2, y + 71, 40, 40);
                        string altYazi = "TEKKAN   F:016477";
                        ffff.Graphics.DrawString(altYazi, myFont, sbrush, x + (i * 68) - 5, y + 2, yaziFormati);



                        ffff.Graphics.DrawString("PD:" + dataGridView1[3, 0].Value.ToString(), myFont, sbrush, x + 33, y + 2, yaziFormati); // tarihi yazdırıyoruz
                        ffff.Graphics.DrawString(dataGridView1[4, 0].Value.ToString(), myFont, sbrush, x + 33, y + 56, yaziFormati);// saati yazdırıyoruz
                        ffff.Graphics.DrawString("P/N:" + dataGridView1[1, 0].Value.ToString() + "-M", myFont, sbrush, x + 20, y + 2, yaziFormati);//stok kodunu yazdırıyoruz
                        ffff.Graphics.DrawString(dataGridView1[2, 0].Value.ToString(), myFont, sbrush, x + 7, y + 2, yaziFormati);// operatörü yazdırıyoruz
                        ffff.Graphics.DrawString("/" + dataGridView1[5, 0].Value.ToString(), myFont, sbrush, x + 7, y + 29, yaziFormati); // Makine adını yazdırıyoruz


                        //x += 68; // yeni barkod için aradaki boşluk
                        ffff.Graphics.DrawString("PD:" + dataGridView1[3, 1].Value.ToString(), myFont, sbrush, x + 101, y + 2, yaziFormati); // tarihi yazdırıyoruz
                        ffff.Graphics.DrawString(dataGridView1[4, 1].Value.ToString(), myFont, sbrush, x + 101, y + 56, yaziFormati); // saati yazdırıyoruz
                        ffff.Graphics.DrawString("P/N:" + dataGridView1[1, 1].Value.ToString() + "-M", myFont, sbrush, x + 88, y + 2, yaziFormati);//stok kodunu yazdırıyoruz
                        ffff.Graphics.DrawString(dataGridView1[2, 1].Value.ToString(), myFont, sbrush, x + 75, y + 2, yaziFormati);//operatörü yazdırıyoruz
                        ffff.Graphics.DrawString("/" + dataGridView1[5, 1].Value.ToString(), myFont, sbrush, x + 75, y + 29, yaziFormati); // Makine adını yazdırıyoruz
                        //x += 68; // yeni barkod için aradaki boşluk

                        ffff.Graphics.DrawString("PD:" + dataGridView1[3, 2].Value.ToString(), myFont, sbrush, x + 169, y + 2, yaziFormati); // tarihi yazdırıyoruz
                        ffff.Graphics.DrawString(dataGridView1[4, 2].Value.ToString(), myFont, sbrush, x + 169, y + 56, yaziFormati); // saati yazdırıyoruz
                        ffff.Graphics.DrawString("P/N:" + dataGridView1[1, 2].Value.ToString() + "-M", myFont, sbrush, x + 156, y + 2, yaziFormati);//stok kodunu yazdırıyoruz
                        ffff.Graphics.DrawString(dataGridView1[2, 2].Value.ToString(), myFont, sbrush, x + 143, y + 2, yaziFormati);// operatörü yazdırıyoruz
                        ffff.Graphics.DrawString("/" + dataGridView1[5, 2].Value.ToString(), myFont, sbrush, x + 143, y + 29, yaziFormati); // Makine adını yazdırıyoruz
                        //x += 68; // yeni barkod için aradaki boşluk

                        ffff.Graphics.DrawString("PD:" + dataGridView1[3, 3].Value.ToString(), myFont, sbrush, x + 237, y + 2, yaziFormati); // tarihi yazdırıyoruz
                        ffff.Graphics.DrawString(dataGridView1[4, 3].Value.ToString(), myFont, sbrush, x + 237, y + 56, yaziFormati); // saati yazdırıyoruz
                        ffff.Graphics.DrawString("P/N:" + dataGridView1[1, 3].Value.ToString() + "-M", myFont, sbrush, x + 224, y + 2, yaziFormati);//stok kodunu yazdırıyoruz
                        ffff.Graphics.DrawString(dataGridView1[2, 3].Value.ToString(), myFont, sbrush, x + 211, y + 2, yaziFormati);// operatörü yazdırıyoruz
                        ffff.Graphics.DrawString("/" + dataGridView1[5, 3].Value.ToString(), myFont, sbrush, x + 211, y + 29, yaziFormati); // Makine adını yazdırıyoruz

                        //x += 68; // yeni barkod için aradaki boşluk
                        ffff.Graphics.DrawString("PD:" + dataGridView1[3, 4].Value.ToString(), myFont, sbrush, x + 309, y + 2, yaziFormati); // tarihi yazdırıyoruz
                        ffff.Graphics.DrawString(dataGridView1[4, 4].Value.ToString(), myFont, sbrush, x + 309, y + 56, yaziFormati); // saati yazdırıyoruz
                        ffff.Graphics.DrawString("P/N:" + dataGridView1[1, 4].Value.ToString() + "-M", myFont, sbrush, x + 296, y + 2, yaziFormati);//stok kodunu yazdırıyoruz
                        ffff.Graphics.DrawString(dataGridView1[2, 4].Value.ToString(), myFont, sbrush, x + 283, y + 2, yaziFormati);// operatörü yazdırıyoruz
                        ffff.Graphics.DrawString("/" + dataGridView1[5, 4].Value.ToString(), myFont, sbrush, x + 283, y + 29, yaziFormati); // Makine adını yazdırıyoruz
                        //x += 68; // yeni barkod için aradaki boşluk

                    }

                    y += 130; // y ekseninde (dikeyde boşluk veriyoruz)

                    //burada sayfa yüksekliğine göre yeni sayfa oluşturuluyor.
                    if (y >= ffff.PageSettings.PrintableArea.Height + 100)
                    {
                        ffff.HasMorePages = true;
                        y = 100;
                        return; // bu fonksiyonun tekrar çalışması için return e ihtiyaç var. yeniden çalıştığında sayac hep 0 dan başlayacağı için global e taşıdık. 
                    }
                    else
                    {
                        ffff.HasMorePages = false;
                    }
                }
                //buraya kadar

                //  MessageBox.Show("Yazdırma Bitti");
                //******************************************************************************************************************************************************
            }
            catch
            {
                MessageBox.Show("Hata");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            sayac = 0;
            printDocument1.Print();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrWhiteSpace(textBox1.Text) ||
                string.IsNullOrWhiteSpace(textBox2.Text) ||
                string.IsNullOrWhiteSpace(textBox4.Text))
            {
                MessageBox.Show("Stok Kodu, Operatör Kodu ve Makine Adı alanlarını doldurmalısınız!");
                return;
            }


            if (checkBox1.Checked == true)
            {
                dataGridView1.ColumnCount = 6;
                dataGridView1.ColumnHeadersVisible = true;

                // Sütun başlığına ait style tanımlıyoruz.
                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();

                columnHeaderStyle.BackColor = Color.Beige;
                columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
                dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

                // Sütun isimlerini giriyoruz.
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                dataGridView1.Columns[0].Name = "SN";
                dataGridView1.Columns[1].Name = "Stok Kodu";
                dataGridView1.Columns[2].Name = "Operatör Kodu";
                dataGridView1.Columns[3].Name = "Tarih";
                dataGridView1.Columns[4].Name = "Saat";
                dataGridView1.Columns[5].Name = "Makine Adı";

                // Satır sayısını tanımlıyoruz

                for (int i = 0; i < Convert.ToInt32(comboBox1.Text); i++)
                {
                    i = dataGridView1.Rows.Add();
                    // 1.satıra verileri yazdırıyoruz
                    dataGridView1.Rows[i].Cells[0].Value = i + 1;
                    dataGridView1.Rows[i].Cells[1].Value = textBox1.Text;
                    dataGridView1.Rows[i].Cells[2].Value = textBox2.Text;
                    dataGridView1.Rows[i].Cells[3].Value = DateTime.Now.ToShortDateString();
                    dataGridView1.Rows[i].Cells[4].Value = DateTime.Now.ToShortTimeString();
                    dataGridView1.Rows[i].Cells[5].Value = textBox4.Text;
                }
            }
            else
            {
                MessageBox.Show("Onaylamadığınız etiketi yazdıramazsınız!");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox3.Text = DateTime.Now.ToShortTimeString();
            comboBox1.SelectedIndex = 0;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            textBox1.Clear();
            textBox2.Clear();
            textBox4.Clear();
            dataGridView1.Refresh();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            excel.Visible = true;

            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);

            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    try
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                        myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    }
                    catch
                    {
                        ;
                    }
                }

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("www.tekkan.com.tr");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            sayac = 0;
            PrintPreviewDialog onizleme = new PrintPreviewDialog();
            onizleme.Document = printDocument1;
            onizleme.ShowDialog();
        }
    }
}


