using System.Collections;
using System.Data.SqlClient;
using System.Drawing;
using Excel= Microsoft.Office.Interop.Excel;

namespace ExcelVTEntegrasyon
{
    public partial class Form1 : Form
    {
        SqlConnection baglanti = new SqlConnection(@"Data Source = OZGUR\SQLEXPRESS; Initial Catalog = projelervt; Integrated Security = True;");
        public Form1()
        {
            InitializeComponent();
        }

        private void BtnVTdenOku_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet sayfa1 = workbook.Sheets[1];  //sayfa1=Sheet1

            string[] basliklar = { "Personel No", "Ad", "Soyad", "�l", "�l�e" };
            Excel.Range range;
            for (int i = 0; i < basliklar.Length; i++)
            {
                range = sayfa1.Cells[1, (1 + i)];
                range.Value2 = basliklar[i];
            }

            try
            {
                baglanti.Open();
                string sorgu = "SELECT PersonelNo, Ad, Soyad, �l, �lce FROM Personel";
                SqlCommand komut = new SqlCommand(sorgu, baglanti);
                SqlDataReader oku = komut.ExecuteReader();

                int sayac = 2;  //ilk sat�r ba�l�klar oldu�u i�in 2. sat�rdan ba�layacak
                while (oku.Read())
                {
                    string PersonelNo = oku[0].ToString();
                    string Ad = oku[1].ToString();
                    string Soyad = oku[2].ToString();
                    string �l = oku[3].ToString();
                    string �lce = oku[4].ToString();
                    richTextBox1.Text = richTextBox1.Text + PersonelNo + "  " + Ad + "  " + Soyad + "  " + �l + "  " + �lce + "\n";
                    range = sayfa1.Cells[sayac, 1];
                    range.Value2 = PersonelNo;
                    range = sayfa1.Cells[sayac, 2];
                    range.Value2 = Ad;
                    range = sayfa1.Cells[sayac, 3];
                    range.Value2 = Soyad;
                    range = sayfa1.Cells[sayac, 4];
                    range.Value2 = �l;
                    range = sayfa1.Cells[sayac, 5];
                    range.Value2 = �lce;
                    sayac++;

                }
            }
            catch (Exception hata)
            {
                MessageBox.Show("SQL Query S�ras�nda Bir Hata Olu�tu. Hata Kodu: SQLREAD01 \n" + hata.ToString());
            }
            finally
            {
                if (baglanti != null)
                    baglanti.Close();
            }
        }

        private void BtnExceldenOku_Click(object sender, EventArgs e)
        {
            Excel.Application excel1;
            Excel.Workbook workbook1;
            Excel.Worksheet worksheet1;
            Excel.Range range;
            int rowCnt = 0;
            int columnCnt = 0;
            excel1 = new Excel.Application();
            workbook1 = excel1.Workbooks.Open(@"C:\Users\ozgur\Downloads\VTExcel.xlsx");
            worksheet1 = (Excel.Worksheet)workbook1.Worksheets.get_Item(1);
            range = worksheet1.UsedRange;  //usedrange, i�inde veri olan t�m h�creleri al�r

            //ilk olarak richTextBox2 nin i�eri�ini temizledim
            richTextBox2.Clear();

            //ilk sat�r ba�l�klar� i�erdi�i i�in 2. sat�rdan ba�layarak verileri al�yorum
            //e�er ilk sat�rda ba�l�klar yoksa 1. sat�rdan ba�layarak verileri alabiliriz

            for (rowCnt = 2; rowCnt <= range.Rows.Count; rowCnt++)
            {
                ArrayList liste = new ArrayList();
                for(columnCnt = 1; columnCnt <= range.Columns.Count; columnCnt++)
                {
                    string okunanhucre = Convert.ToString((range.Cells[rowCnt, columnCnt]as Excel.Range).Value2);
                    richTextBox2.Text = richTextBox2.Text + okunanhucre + "  ";
                    liste.Add(okunanhucre);
                }
                richTextBox2.Text = richTextBox2.Text + "\n";

                //liste i�indeki verileri veritaban�na yazd�rma i�lemi
                try
                {
                    baglanti.Open();
                    string sorgu = "INSERT INTO Personel(PersonelNo, Ad, Soyad, �l, �lce) VALUES(@PersonelNo, @Ad, @Soyad, @�l, @�lce)";
                    SqlCommand komut = new SqlCommand(sorgu, baglanti);
                    komut.Parameters.AddWithValue("@PersonelNo", liste[0]);
                    komut.Parameters.AddWithValue("@Ad", liste[1]);
                    komut.Parameters.AddWithValue("@Soyad", liste[2]);
                    komut.Parameters.AddWithValue("@�l", liste[3]);
                    komut.Parameters.AddWithValue("@�lce", liste[4]);
                    komut.ExecuteNonQuery();
                }
                catch (Exception hata)
                {
                    MessageBox.Show("SQL Veri Taban�na Yazma S�ras�nda Bir Hata Olu�tu. Hata Kodu: SQLWRITE01 \n" + hata.ToString());
                }
                finally
                {
                    if (baglanti != null)
                        baglanti.Close();
                }
            }
            excel1.Quit();
            ReleaseObject(worksheet1);
            ReleaseObject(workbook1);
            ReleaseObject(excel1);
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch(Exception hata)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
