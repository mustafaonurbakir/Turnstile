using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Management.Instrumentation;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace deneme
{

    public partial class Turnike : Form
    {
        public Turnike()
        {
            InitializeComponent();
            label1.Text = "Dosya bekleniyor";
            comboBox1.SelectedIndex = 0;
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        DialogResult dialogresult;
        string sicil_str, ad_str, soyad_str, tarih_str, hareket_str, turnike_str, bolum_str;


        //dosya seçme butonu
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    listView3.Items[0].ImageIndex = 0; //status
                    listView3.Items[1].ImageIndex = 1;
                    listView3.Items[2].ImageIndex = 1;
                    listView3.Items[3].ImageIndex = 1;
                    listView3.Items[4].ImageIndex = 1;
                    listView3.Items[5].ImageIndex = 1;
                    listView3.Items[5].Text = "Bitti";
                    listView2.Items.Clear();

                    label1.Text = "Dosya analiz ediliyor.";//status

                    dialogresult = DialogResult.OK;
                    var sheetname = GetExcelSheetNames(openFileDialog1.FileName)[0];
                    OleDbConnection xlsxbaglanti = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog1.FileName + "; Extended Properties='Excel 12.0 Xml;HDR=YES'");
                    System.Data.DataTable tablo = new System.Data.DataTable();
                    
                    try
                    {
                        xlsxbaglanti.Open();
                        tablo.Clear();
                        OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + sheetname + "]", xlsxbaglanti);
                        da.Fill(tablo);
                        dataGridView1.DataSource = tablo;
                        foreach (DataColumn row in tablo.Columns)
                        {
                            Console.WriteLine(row.ColumnName);
                            if(row.ColumnName.ToLower() == "ad" || row.ColumnName.ToLower() == "isim")
                            {
                                ad_str = row.ColumnName;
                            }
                            else if(comboBox1.SelectedIndex == 0 && (row.ColumnName.ToLower() == "soyad"))
                            {
                                soyad_str = row.ColumnName;
                            }
                            else if(row.ColumnName.ToLower() == "tarih")
                            {
                                tarih_str = row.ColumnName;
                            }
                            else if(row.ColumnName.ToLower() == "hareket" || row.ColumnName.ToLower() == "giriş-çıkış")
                            {
                                hareket_str = row.ColumnName;
                            }
                            else if(row.ColumnName.ToLower() == "bölüm" || row.ColumnName.ToLower() == "departman" || row.ColumnName.ToLower() == "bölümü" || row.ColumnName.ToLower() == "bölüm adı")
                            {
                                bolum_str = row.ColumnName;
                            }
                            else if(row.ColumnName.ToLower() == "sicil" || row.ColumnName.ToLower() == "sicilno")
                            {
                                sicil_str = row.ColumnName;
                            }
                            else if(row.ColumnName.ToLower() == "kapı adı")
                            {
                                turnike_str = row.ColumnName;
                            }
                            else
                            {
                                Tanılama newform = new Tanılama(row.ColumnName);
                                newform.ShowDialog();
                            
                                string secilendeger = newform.ReturnValue1;
                                if (secilendeger == "Sicil")
                                {
                                    sicil_str = row.ColumnName;
                                }
                                else if (secilendeger == "Ad")
                                {
                                    ad_str = row.ColumnName;
                                }
                                else if (secilendeger == "Soyad")
                                {
                                    soyad_str = row.ColumnName;
                                }
                                else if (secilendeger == "Tarih")
                                {
                                    tarih_str = row.ColumnName;
                                }
                                else if (secilendeger == "Hareket(Giris / Çıkış)")
                                {
                                    hareket_str = row.ColumnName;
                                }
                                else if (secilendeger == "Turnike Adı")
                                {
                                    turnike_str = row.ColumnName;
                                }
                                else if (secilendeger == "Departman")
                                {
                                    bolum_str = row.ColumnName;
                                }
                                else if (secilendeger == "Diğer")
                                {

                                }
                                else MessageBox.Show("Seçilen değer bulunamadı!");
                            }
                        }
                        xlsxbaglanti.Close();

                        listView3.Items[1].ImageIndex = 0;
                        label1.Text = "'Do It' basınız";

                        button1.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Programda Hata Meydana Geldi." + Environment.NewLine + "Hata : " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    finally
                    {
                        if (xlsxbaglanti.State == ConnectionState.Open)
                        {
                            xlsxbaglanti.Close();
                        }
                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Form1", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private String[] GetExcelSheetNames(string excelFile)
        {
            String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFile + "; Extended Properties='Excel 12.0 Xml;HDR=YES'";
            OleDbConnection objConn = new OleDbConnection(connString);
            System.Data.DataTable dt = null;

            try
            {
                // Connection String. Change the excel file to the file you
                // will search.
                // Create connection object by using the preceding connection string.
                // Open connection with the database.
                objConn.Open();
                // Get the data table containg the schema guid.
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                // Loop through all of the sheets if you want too...
                for (int j = 0; j < excelSheets.Length; j++)
                {
                    // Query each excel sheet.
                }

                return excelSheets;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void hakkındaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Hakkında().Show();
        }

        private void çıkışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void yardımToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            new Yardım().Show();
        }


        //do it butonu
        private void button1_Click(object sender, EventArgs e)
        {
            if (dialogresult == DialogResult.OK)
            {
                label1.Text = "Dosya okunuyor";//status
                var sheetname = GetExcelSheetNames(openFileDialog1.FileName)[0];

                List<Person> ListPerson = new List<Person>();
                DateTime firstdate = DateTime.Now, lastdate = DateTime.Now;

                OleDbConnection xlsxbaglanti = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog1.FileName + "; Extended Properties='Excel 12.0 Xml;HDR=YES'");
                    
                //OKUMA
                try
                {
                    xlsxbaglanti.Open();

                    OleDbCommand komut = new OleDbCommand("SELECT * FROM [" + sheetname + "]", xlsxbaglanti);
                    OleDbDataReader oku = komut.ExecuteReader();

                    //variables
                    Person tempperson = null;
                    int temppersonnumber = 0;
                    bool newperson = true;

                    while (oku.Read())
                    {
                        newperson = true;

                        //dosyadan okuma
                        string ad = oku[ad_str].ToString();
                        string soyad = oku[soyad_str].ToString();
                        string adSoyad;
                        if (comboBox1.SelectedIndex == 0)   adSoyad = ad + " " + soyad;
                        else adSoyad = ad;
                        string departman = oku[bolum_str].ToString();
                        string mydate = oku[tarih_str].ToString();
                        DateTime Date;
                        if (mydate[1].ToString() == "." && mydate[3].ToString() == ".")
                        {
                            Date = DateTime.ParseExact(mydate, "d.M.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else if (mydate[1].ToString() == ".")
                        {
                            Date = DateTime.ParseExact(mydate, "d.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else if (mydate[4].ToString() == ".")
                        {
                            Date = DateTime.ParseExact(mydate, "dd.M.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            Date = DateTime.ParseExact(mydate, "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        }
                        string hareket = oku[hareket_str].ToString();
                        string turnike = oku[turnike_str].ToString();

                        //hareket tanıma
                        hareket = hareket.ToLower();
                        if (hareket == "giris" || hareket == "giriş" || hareket == "gırıs" || hareket == "gırış")
                        {
                            hareket = "g";
                        }
                        else if (hareket == "çıkış" || hareket == "cıkış" || hareket == "çikiş" || hareket == "cikiş"
                            || hareket == "cıkıs" || hareket == "cikis" || hareket == "çikis" || hareket == "çikiş")
                        {
                            hareket = "c";
                        }
                        else MessageBox.Show("Giriş/Çıkış hareketi ayırt edilemiyor!");

                        //insan var mı yok mu?
                        if (ListPerson.Count() != 0)
                        {
                            for (int i = 0; i < ListPerson.Count(); i++)
                            {
                                if (ListPerson[i].Name == adSoyad)
                                {
                                    newperson = false;
                                    tempperson = ListPerson[i];
                                    temppersonnumber = i;
                                }
                            }
                        }

                        //insan yoksa ekleme, varsa hesaplama
                        if (newperson && hareket == "g")
                        {
                            //firstdate ataması
                            if (ListPerson.Count() == 0)
                            {
                                firstdate = Date;
                            }

                            //insan ekleme
                            Person newPerson = new Person
                            {
                                Name = adSoyad,
                                Date = Date,
                                Departman = departman,
                                lastmove = "g",
                                invalid_entry = 0,
                                invalid_out = 0
                            };
                            newPerson.girissayisi.Add(Date.Month.ToString() + Date.Year.ToString(), 1);

                            ListPerson.Add(newPerson);
                        }
                        else if (!newperson)
                        {
                            //yemek turnikesi
                            if (turnike.Substring(0, 5).ToLower() == "yemek")
                            {
                                string[] row = { "Yemek turnikesi giriş olarak kaydedilmiş!", ListPerson[temppersonnumber].Name, ListPerson[temppersonnumber].Date.ToString(), null };
                                var ListViewItem = new ListViewItem(row);
                                listView2.Items.Add(ListViewItem);
                                listView2.Items[listView2.Items.Count - 1].Group = listView2.Groups[2];
                            }
                            //çıkış
                            else if (tempperson.lastmove == "g" && hareket == "c")
                            {
                                if (Date.Day == tempperson.Date.Day && Date.Month == tempperson.Date.Month)
                                {
                                    lastdate = Date;
                                    TimeSpan timedifferance = Date - tempperson.Date;

                                    if (!ListPerson[temppersonnumber].in_time_day.ContainsKey(tempperson.Date.Date))
                                    {
                                        ListPerson[temppersonnumber].in_time_day.Add(tempperson.Date.Date, 0);
                                    }
                                    ListPerson[temppersonnumber].in_time_day[tempperson.Date.Date] += timedifferance.TotalMinutes;
                                    if (!ListPerson[temppersonnumber].cikissayisi.ContainsKey(Date.Month.ToString() + Date.Year.ToString()))
                                    {
                                        ListPerson[temppersonnumber].cikissayisi.Add(Date.Month.ToString() + Date.Year.ToString(), 0);
                                    }
                                    ListPerson[temppersonnumber].cikissayisi[Date.Month.ToString() + Date.Year.ToString()] += 1;
                                    ListPerson[temppersonnumber].lastmove = "c";
                                    ListPerson[temppersonnumber].Date = Date;
                                }
                                else
                                {
                                    string[] row = { "Tarihleri arasında içeride kaldı!", ListPerson[temppersonnumber].Name, ListPerson[temppersonnumber].Date.ToString(), Date.ToString() };
                                    var ListViewItem = new ListViewItem(row);
                                    listView2.Items.Add(ListViewItem);
                                    listView2.Items[listView2.Items.Count - 1].Group = listView2.Groups[3];

                                    lastdate = Date;
                                    TimeSpan timedifferance = Date - tempperson.Date;

                                    DateTime temp_tempperson = tempperson.Date;
                                    TimeSpan newtimedifferance = tempperson.Date.AddDays(1).Date - temp_tempperson;
                                    for(;temp_tempperson.Date.Day != Date.AddDays(1).Date.Day;)
                                    {
                                        if (!ListPerson[temppersonnumber].in_time_day.ContainsKey(temp_tempperson.Date))
                                        {
                                            ListPerson[temppersonnumber].in_time_day.Add(temp_tempperson.Date, 0);
                                        }
                                        ListPerson[temppersonnumber].in_time_day[temp_tempperson.Date] += newtimedifferance.TotalMinutes;

                                        temp_tempperson = temp_tempperson.AddDays(1);
                                        timedifferance = timedifferance - newtimedifferance;
                                        if (timedifferance.TotalHours > 24) newtimedifferance = TimeSpan.FromDays(1);
                                        else newtimedifferance = timedifferance;
                                    }
                                    
                                    
                                    if (!ListPerson[temppersonnumber].cikissayisi.ContainsKey(Date.Month.ToString() + Date.Year.ToString()))
                                    {
                                        ListPerson[temppersonnumber].cikissayisi.Add(Date.Month.ToString() + Date.Year.ToString(), 0);
                                    }
                                    ListPerson[temppersonnumber].cikissayisi[Date.Month.ToString() + Date.Year.ToString()] += 1;

                                    ListPerson[temppersonnumber].lastmove = "c";
                                    ListPerson[temppersonnumber].Date = Date;
                                }

                            }
                            //giriş
                            else if (tempperson.lastmove == "c" && hareket == "g")
                            {
                                if (!ListPerson[temppersonnumber].girissayisi.ContainsKey(Date.Month.ToString() + Date.Year.ToString()))
                                {
                                    ListPerson[temppersonnumber].girissayisi.Add(Date.Month.ToString() + Date.Year.ToString(), 0);
                                }
                                ListPerson[temppersonnumber].girissayisi[Date.Month.ToString() + Date.Year.ToString()] += 1;
                                ListPerson[temppersonnumber].lastmove = "g";
                                ListPerson[temppersonnumber].Date = Date;
                            }
                            //çift çıkış ERROR
                            else if (tempperson.lastmove == "c" && hareket == "c")
                            {
                                if (!ListPerson[temppersonnumber].gecersiz_cikis.ContainsKey(Date.Month.ToString() + Date.Year.ToString()))
                                {
                                    ListPerson[temppersonnumber].gecersiz_cikis.Add(Date.Month.ToString() + Date.Year.ToString(), 0);
                                }
                                ListPerson[temppersonnumber].gecersiz_cikis[Date.Month.ToString() + Date.Year.ToString()] += 1;
                                ListPerson[temppersonnumber].invalid_out += 1;
                                string[] row = { "Art arda iki çıkış oldu!", ListPerson[temppersonnumber].Name, ListPerson[temppersonnumber].Date.ToString(), Date.ToString() };
                                var ListViewItem = new ListViewItem(row);
                                listView2.Items.Add(ListViewItem);
                                listView2.Items[listView2.Items.Count - 1].Group = listView2.Groups[1];
                            }
                            //çift giriş ERROR
                            else if (tempperson.lastmove == "g" && hareket == "g")
                            {
                                if (!ListPerson[temppersonnumber].gecersiz_giris.ContainsKey(Date.Month.ToString() + Date.Year.ToString()))
                                {
                                    ListPerson[temppersonnumber].gecersiz_giris.Add(Date.Month.ToString() + Date.Year.ToString(), 0);
                                }
                                ListPerson[temppersonnumber].gecersiz_giris[Date.Month.ToString() + Date.Year.ToString()] += 1;
                                ListPerson[temppersonnumber].invalid_entry += 1;
                                ListPerson[temppersonnumber].invalid_out += 1;
                                string[] row = { "Art arda iki giriş oldu!", ListPerson[temppersonnumber].Name, ListPerson[temppersonnumber].Date.ToString(), Date.ToString() };
                                var ListViewItem = new ListViewItem(row);
                                listView2.Items.Add(ListViewItem);
                                listView2.Items[listView2.Items.Count - 1].Group = listView2.Groups[0];
                            }
                            else MessageBox.Show("Garip bir hata meydana geldi!"+ Environment.NewLine+ "Giris/cikis arasında tanımlanmayan durum. Line:371");

                        }

                    }
                    listView3.Items[2].ImageIndex = 0;//status
                    xlsxbaglanti.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Programda Hata Meydana Geldi." + Environment.NewLine + "Hata : " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                finally
                {
                    if(xlsxbaglanti.State == ConnectionState.Open)
                    {
                        xlsxbaglanti.Close();
                    }
                }

                //YAZMA
                try
                {
                    label1.Text = "dosyaya yazılıyor";//status

                    //Yeni excel belgesi yaratılıyor
                    Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }
                    var xlWorkBook = xlApp.Workbooks.Add();

                    //Yeni worksheet oluşturuluyor
                    var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    listView3.Items[3].ImageIndex = 0;

                    //Sütunlar oluşturuluyor
                    xlWorkSheet.Cells[1, 1] = "Ad Soyad";
                    xlWorkSheet.Cells[1, 2] = "Departman";
                    DateTime temp_firstdate = firstdate;
                    int j = 0;
                    for (j = 3; temp_firstdate.Date.Month != lastdate.Date.AddMonths(1).Month; j++)
                    {
                        xlWorkSheet.Cells[1, j] = temp_firstdate.Date.ToString("MMMM");j++;
                        xlWorkSheet.Cells[1, j] = "Giriş";j++;
                        xlWorkSheet.Cells[1, j] = "Çıkış";j++;
                        xlWorkSheet.Cells[1, j] = "Geçersiz Giriş";j++;
                        xlWorkSheet.Cells[1, j] = "Gerçersiz Çıkış";
                        temp_firstdate = temp_firstdate.AddMonths(1);
                    }
                    temp_firstdate = firstdate;
                    for (; temp_firstdate.Date != lastdate.Date.AddDays(1); j++)
                    {
                        xlWorkSheet.Cells[1, j] = temp_firstdate.Date;
                        temp_firstdate = temp_firstdate.AddDays(1);
                    }

                    //Veriler yazılıyor
                    for (int i = 0; i < ListPerson.Count(); i++)
                    {
                        xlWorkSheet.Cells[i + 2, 1] = ListPerson[i].Name;
                        xlWorkSheet.Cells[i + 2, 2] = ListPerson[i].Departman;
                        
                        //aylık
                        temp_firstdate = firstdate;
                        for (j = 3; temp_firstdate.Date.Month != lastdate.Date.AddMonths(1).Month; j++)
                        {
                            double temptotal = 0;

                            DateTime temp_firstdate2 = firstdate;
                            for (int k = 3; temp_firstdate2.Date != lastdate.Date.AddDays(1); k++)
                            {
                                if (ListPerson[i].in_time_day.ContainsKey(temp_firstdate2.Date.Date))
                                {
                                    if (temp_firstdate2.Date.Month == temp_firstdate.Date.Month)
                                        temptotal += ListPerson[i].in_time_day[temp_firstdate2.Date];
                                    else temptotal += 0;
                                }

                                temp_firstdate2 = temp_firstdate2.AddDays(1);
                            }
                            xlWorkSheet.Cells[i + 2, j] = (int)temptotal/60 + ":" + (int)temptotal%60; j++;

                            if (ListPerson[i].girissayisi.ContainsKey(temp_firstdate.Month.ToString() + temp_firstdate.Year.ToString()))
                            {
                                xlWorkSheet.Cells[i + 2, j] = ListPerson[i].girissayisi[temp_firstdate.Month.ToString() + temp_firstdate.Year.ToString()]; j++;
                            }
                            else
                            {
                                xlWorkSheet.Cells[i + 2, j] = 0;
                                j++;
                            }
                            if (ListPerson[i].cikissayisi.ContainsKey(temp_firstdate.Month.ToString() + temp_firstdate.Year.ToString()))
                            {
                                xlWorkSheet.Cells[i + 2, j] = ListPerson[i].cikissayisi[temp_firstdate.Month.ToString() + temp_firstdate.Year.ToString()]; j++;
                            }
                            else
                            {
                                xlWorkSheet.Cells[i + 2, j] = 0;
                                j++;
                            }
                            if (ListPerson[i].gecersiz_giris.ContainsKey(temp_firstdate.Month.ToString() + temp_firstdate.Year.ToString()))
                            {
                                xlWorkSheet.Cells[i + 2, j] = ListPerson[i].gecersiz_giris[temp_firstdate.Month.ToString() + temp_firstdate.Year.ToString()]; j++;
                            }
                            else
                            {
                                xlWorkSheet.Cells[i + 2, j] = 0;
                                j++;
                            }
                            if (ListPerson[i].gecersiz_cikis.ContainsKey(temp_firstdate.Month.ToString() + temp_firstdate.Year.ToString()))
                            {
                                xlWorkSheet.Cells[i + 2, j] = ListPerson[i].gecersiz_cikis[temp_firstdate.Month.ToString() + temp_firstdate.Year.ToString()];
                            }
                            else xlWorkSheet.Cells[i + 2, j] = 0; 

                            temp_firstdate = temp_firstdate.AddMonths(1);
                        }
                        
                        //günlük
                        temp_firstdate = firstdate;
                        for (; temp_firstdate.Date != lastdate.Date.AddDays(1); j++)
                        {
                            if (ListPerson[i].in_time_day.ContainsKey(temp_firstdate.Date.Date))
                                xlWorkSheet.Cells[i + 2, j] = (int)ListPerson[i].in_time_day[temp_firstdate.Date.Date]/60 + ":" + (int)ListPerson[i].in_time_day[temp_firstdate.Date.Date]%60;
                            else xlWorkSheet.Cells[i + 2, j] = 0;
                            temp_firstdate = temp_firstdate.AddDays(1);
                        }
                    }
                    label1.Text = "dosyaya yazılıyor son aşama";
                    listView3.Items[4].ImageIndex = 0;

                    //excel kayıt ediliyor
                    string file_name = System.IO.Path.GetDirectoryName(openFileDialog1.FileName) + "\\Output " + openFileDialog1.SafeFileName;
                    //label6.Text = file_name;
                    xlWorkBook.SaveAs(file_name);
                    xlWorkBook.Close();
                    listView3.Items[listView3.Items.Count - 1].Text = "Bitti - File path:" + file_name;
                    label1.Text = "Bitti";
                    listView3.Items[5].ImageIndex = 0;
                }
                catch (Exception ex)
                {
                    label1.Text = "dosyaya yazmada hata oluştu!";
                    MessageBox.Show("Yazma sırasında bir hata meydana geldi" + Environment.NewLine + "Hata : " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else MessageBox.Show("Önce excel dosyası seçmen gerekiyor!", "Ao!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            
        }
    }

}
