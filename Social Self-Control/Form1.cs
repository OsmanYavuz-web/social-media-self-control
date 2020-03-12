using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// eklenenler
using System.IO;
using System.Net;
using System.Threading;
using HtmlAgilityPack;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using TweetSharp;

namespace Social_Self_Control
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        // Değişkenker
        string yt_KaynakKod = null;
        bool girisDurumu = false;
        string hesapAdı = null;
        string kaynakKod = null;
        Thread t_VideoCek;
        int sayfa = 0;
        int satirSayisi = 0;
        bool girisIslemi = false;
        bool paylasimMotoru = false;



        #region Veri Ayıklama Fonksiyon
        public string ayiklananVeri;
        void veriAyiklama(string kaynakKod, string ilkVeri, int ilkVeriKS, string sonVeri)
        {
            try
            {
                string gelen = kaynakKod;
                int titleIndexBaslangici = gelen.IndexOf(ilkVeri) + ilkVeriKS;
                int titleIndexBitisi = gelen.Substring(titleIndexBaslangici).IndexOf(sonVeri);
                ayiklananVeri = gelen.Substring(titleIndexBaslangici, titleIndexBitisi);
            }
            catch //(Exception ex)
            {
                //MessageBox.Show("Hata: " + ex.Message, "Hata;", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region SefLink Fonksiyon
        public string sefLink(string link)
        {
            string don = link.ToLower();
            don = don.Replace("ş", "s").Replace("ı", "i").Replace("ğ", "g").Replace("ç", "c").Replace(".", "").Replace(",", "").Replace(" ", "-");
            don = don.Replace("+", "plus").Replace("#", "sharp").Replace("ü", "u").Replace("ö", "o").Replace("?", "");
            return don;
        }
        #endregion

        #region FORM_LOAD
        private void Form1_Load(object sender, EventArgs e)
        {
            //Thread Çalıştırma
            CheckForIllegalCrossThreadCalls = false;
        }
        #endregion

        #region FORM_SHOWN - Form açıldığında
        private void Form1_Shown(object sender, EventArgs e)
        {
            // Bilgi Mesajı
            statusLabel.ForeColor = Color.Blue;
            statusLabel.Text = "Program başlatıldı.";
            listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Program başlatıldı.");

            // Bilgi Mesajı 
            statusLabel.ForeColor = Color.Blue;
            statusLabel.Text = "Youtube'a bağlanılıyor. Lütfen bekleyin..";
            listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Youtube'a bağlanılıyor. Lütfen bekleyin..");

        }
        #endregion

        #region FORM_CLOSED & FORM_CLOSING - Form kapandığında
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            // cookie temizleme 
            Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 255");
            //Application.Exit();
            //this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Application.Exit();
            //this.Close();
        }
        #endregion

        #region webBrowser - Youtube Sayfası | webBrowser_DocumentCompleted
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            try
            {
                // kaynak kod aktar
                yt_KaynakKod = webBrowser1.Document.Body.InnerHtml.ToString();
                richTextBox_ytKaynakKod.Text = yt_KaynakKod;

                // webBrowser yüklendiyse
                if (this.webBrowser1.ReadyState != WebBrowserReadyState.Complete)
                {
                    return;
                }
                else
                {
                    #region Profil Açık Kaldıysa - IF
                    // profil açık kalmışsa
                    if (yt_KaynakKod.IndexOf("<div class=\"yt-masthead-picker-name\" dir=\"ltr\">") != -1)
                    {
                        //MessageBox.Show("1");

                        webBrowser1.Navigate("https://accounts.google.com/ServiceLogin?sacu=1&continue=https%3A%2F%2Fm.youtube.com%2Fsignin%3Fapp%3Dm%26action_handle_signin%3Dtrue%26next%3D%252Fmy_videos%253Fclient%253Dmv-google%2526gl%253DTR%2526hl%253Dtr%26feature%3Dmobile%26hl%3Dtr&hl=tr&service=youtube");
                        girisDurumu = false;
                    }
                    #endregion

                    if (hesapAdı == null)
                    {
                        #region Giriş durumu
                        // giriş durumu
                        if (girisDurumu == false)
                        {
                            // nesne aktif
                            groupBox_ytGirisForm.Enabled = true;

                            #region Youtube Anasayfasındaysa - IF
                            // anasayfadaysa
                            if (richTextBox_ytKaynakKod.Text.IndexOf(">Oturum aç</A>") != -1)
                            {
                                webBrowser1.Navigate("https://accounts.google.com/ServiceLogin?passive=true&hl=tr&uilel=3&continue=https%3A%2F%2Fwww.youtube.com%2Fsignin%3Fnext%3D%252F%26action_handle_signin%3Dtrue%26hl%3Dtr%26app%3Ddesktop%26feature%3Dsign_in_button&service=youtube#identifier");
                            }
                            #endregion

                            #region Giriş yap sayfasındaysa - ELSE IF
                            // giriş yap
                            else if (richTextBox_ytKaynakKod.Text.IndexOf("YouTube'a devam etmek için oturum açın") != -1 && richTextBox_ytKaynakKod.Text.IndexOf("E-posta adresinizi girin") != -1)
                            {
                                // nesne durumu
                                button_ytGiris.Enabled = true;
                                textBox_ytEposta.ReadOnly = false;
                                textBox_ytParola.ReadOnly = false;
                                button_ytCikis.Enabled = false;
                                tabControl1.Enabled = true;

                                // Bilgi Mesajı
                                label_ytDurum.Text = "Oturum Açabilirsiniz.";
                                statusLabel.ForeColor = Color.Blue;
                                statusLabel.Text = "Oturum açabilirsiniz.";
                                listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Oturum açabilirsiniz.");
                            }
                            #endregion

                            #region Şifre yanlışsa - ELSE IF
                            // şifre yanlış
                            else if (richTextBox_ytKaynakKod.Text.IndexOf("Şifre yanlış.") != -1)
                            {
                                // Bilgi Mesajı
                                label_ytDurum.Text = "Parola hatalı.";
                                label_ytDurum.Text = "Parola hatalı.";
                                statusLabel.ForeColor = Color.DarkRed;
                                statusLabel.Text = "Parola hatalı.";
                                listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Parola hatalı.");

                            }
                            #endregion

                            #region Giriş Başarılırsa - IF
                            // giriş işlemi başlamışsa
                            if (richTextBox_ytKaynakKod.Text.IndexOf("<LABEL class=stacked-label for=Passwd>Şifre</LABEL>") != -1)
                            {
                                // parola
                                webBrowser1.Document.GetElementById("Passwd").InnerText = textBox_ytParola.Text;

                                // ileri butonu
                                HtmlElementCollection elc2 = webBrowser1.Document.GetElementsByTagName("input");
                                foreach (HtmlElement el2 in elc2)
                                {
                                    if (el2.GetAttribute("name").Equals("signIn"))
                                    {
                                        el2.InvokeMember("click");
                                    }
                                }

                            }
                            #endregion

                            #region Giriş Yapılmışsa
                            // giriş yapılmışsa
                            if (richTextBox_ytKaynakKod.Text.IndexOf(">Hesabım</A></DIV>") != -1 && richTextBox_ytKaynakKod.Text.IndexOf(">Çıkış</A>") != -1)
                            {
                                // giriş durumu aktif
                                girisDurumu = true;

                                // nesne durumu

                                button_ytGiris.Enabled = false;
                                textBox_ytEposta.ReadOnly = true;
                                textBox_ytParola.ReadOnly = true;
                                button_ytCikis.Enabled = true;
                                tabControl1.Enabled = true;

                                girisIslemi = true;


                                // videolar sayfasına git
                                if (richTextBox_ytKaynakKod.Text.IndexOf("Videolar") != -1)
                                {
                                    // Bilgi Mesajı
                                    label_ytDurum.Text = "Oturum Açılıyor.";
                                    statusLabel.ForeColor = Color.DarkGreen;
                                    statusLabel.Text = "Oturum Açılıyor.";
                                    listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Oturum Açılıyor.");

                                    // hesap adı
                                    veriAyiklama(yt_KaynakKod, "<SPAN style=\"FONT-WEIGHT: bold; COLOR: #000000\">", 48, "</SPAN>");
                                    hesapAdı = ayiklananVeri;

                                    // Bilgi Mesajı
                                    label_ytDurum.Text = "Hoşgeldin, " + hesapAdı;
                                    statusLabel.ForeColor = Color.DarkGreen;
                                    statusLabel.Text = "Hoşgeldin, " + hesapAdı;
                                    listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Hoşgeldin, " + hesapAdı);

                                    // yönlendir
                                    webBrowser1.Navigate("https://m.youtube.com/my_videos?gl=TR&hl=tr&client=mv-google");
                                }
                                else
                                {
                                    // yönlendirme
                                    webBrowser1.Navigate("https://m.youtube.com/my_videos?gl=TR&hl=tr&client=mv-google");
                                }
                            }
                            #endregion
                        }
                        else
                        {
                            #region Giriş Yapışmışsa - IF
                            // giriş yapılmışsa
                            if (richTextBox_ytKaynakKod.Text.IndexOf(">Hesabım</A></DIV>") != -1 && richTextBox_ytKaynakKod.Text.IndexOf(">Çıkış</A>") != -1)
                            {
                                // giriş durumu
                                girisDurumu = true;

                                // nesne durumu
                                button_ytGiris.Enabled = false;
                                textBox_ytEposta.ReadOnly = true;
                                textBox_ytParola.ReadOnly = true;
                                button_ytCikis.Enabled = true;

                                // videolar sayfasına git
                                if (richTextBox_ytKaynakKod.Text.IndexOf("Videolar") != -1)
                                {
                                    // Bilgi Mesajı
                                    label_ytDurum.Text = "Oturum Açılıyor.";
                                    statusLabel.ForeColor = Color.DarkGreen;
                                    statusLabel.Text = "Oturum Açılıyor.";
                                    listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Oturum Açılıyor.");

                                    // hesap adı
                                    veriAyiklama(yt_KaynakKod, "<div class=\"yt-masthead-picker-name\" dir=\"ltr\">", 47, "</div>");
                                    hesapAdı = ayiklananVeri;

                                    // Bilgi Mesajı
                                    label_ytDurum.Text = "Hoşgeldin, " + hesapAdı;
                                    statusLabel.ForeColor = Color.DarkGreen;
                                    statusLabel.Text = "Hoşgeldin, " + hesapAdı;
                                    listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Hoşgeldin, " + hesapAdı);

                                    // yönlendir
                                    webBrowser1.Navigate("https://m.youtube.com/my_videos?gl=TR&hl=tr&client=mv-google");
                                }
                                else
                                {
                                    // yönlendir
                                    webBrowser1.Navigate("https://m.youtube.com/my_videos?gl=TR&hl=tr&client=mv-google");
                                }
                            }
                            #endregion
                        }
                        #endregion
                    }
                }

                if (richTextBox_ytKaynakKod.Text.IndexOf("<span class=\"yt-uix-button-content\">Oturum aç</span>") != -1)
                {
                    webBrowser1.Navigate("https://m.youtube.com/?persist_app=1&app=m");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region Giriş Yap Butonu
        private void button_ytGiris_Click(object sender, EventArgs e)
        {
            try
            {
                // eposta
                webBrowser1.Document.GetElementById("Email").InnerText = textBox_ytEposta.Text;

                // ileri butonu
                HtmlElementCollection elc2 = webBrowser1.Document.GetElementsByTagName("input");
                foreach (HtmlElement el2 in elc2)
                {
                    if (el2.GetAttribute("value").Equals("İleri"))
                    {
                        el2.InvokeMember("click");
                    }
                }
            }
            catch
            {
                // Bilgi Mesajı
                MessageBox.Show("Giriş işlemi yapılırken bir sorun oluştu.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                statusLabel.ForeColor = Color.DarkRed;
                statusLabel.Text = "Giriş işlemi yapılırken bir sorun oluştu.";
                listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Giriş işlemi yapılırken bir sorun oluştu.");
            }
        }
        #endregion

        #region Çıkış Yap Butonu
        private void button_ytCikis_Click(object sender, EventArgs e)
        {
            // cookie temizleme 
            Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 255");

            // nesne durumu
            girisDurumu = false;
            tabControl1.Enabled = false;
            girisIslemi = false;

            // Bilgi Mesajı
            label_ytDurum.Text = hesapAdı + " çıkış yapılıyor.";
            statusLabel.ForeColor = Color.DarkGreen;
            statusLabel.Text = hesapAdı + " çıkış yapılıyor, bekleyin.";
            listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] " + hesapAdı + " çıkış yapılıyor, bekleyin.");

            //
            hesapAdı = null;

            // çıkış yap
            webBrowser1.Navigate("https://m.youtube.com/logout?gl=TR&amp;hl=tr&amp;client=mv-google");
            // mobile yönlendir
            //webBrowser1.Navigate("https://m.youtube.com/?persist_app=1&app=m");
        }
        #endregion

        #region Videolarımı Çek Fonksiyonu
        void videolarimiCek()
        {
            // videolarım sayfasında mı ?
            if (richTextBox_ytKaynakKod.Text.IndexOf("Videolarım") != -1)
            {
                sayfa = sayfa + 1;
                webBrowser2.Navigate("https://m.youtube.com/my_videos?gl=TR&client=mv-google&hl=tr&p=" + sayfa);
            }
        }
        #endregion

        #region webBrowser - Youtube Videolarım Sayfası | webBrowser2_DocumentCompleted
        private void webBrowser2_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            kaynakKod = webBrowser2.Document.Body.InnerHtml.ToString();
            richTextBox1.Text = kaynakKod;

            // başka sayfa var mı ?
            if (kaynakKod.IndexOf("Sonraki sayfa »") != -1)
            {
                try
                {
                    // HtmlDocument sınıf tanımlama
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(kaynakKod);

                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//td[@class='videoListItem']");
                    foreach (var veri in XPath)
                    {
                        // veri
                        string video = veri.InnerHtml;

                        // video id
                        veriAyiklama(video, "/watch?v=", 9, "&amp;");
                        string video_id = ayiklananVeri;

                        // video adı
                        veriAyiklama(video, video_id, video_id.Length + 43, " </a>");
                        string video_adi = ayiklananVeri;

                        // Satır sayısını tanımlıyoruz
                        satirSayisi = dataGridView_videoListesi.Rows.Add();

                        // 1.satıra verileri yazdırıyoruz
                        dataGridView_videoListesi.Rows[satirSayisi].Cells[0].Value = satirSayisi + 1;
                        dataGridView_videoListesi.Rows[satirSayisi].Cells[1].Value = video_id;
                        dataGridView_videoListesi.Rows[satirSayisi].Cells[2].Value = video_adi;
                        dataGridView_videoListesi.Rows[satirSayisi].Cells[4].Value = "Paylaşılmadı";
                        dataGridView_videoListesi.Rows[satirSayisi].Cells[3].Value = "";

                        // toplam video
                        label_topVideo.Text = dataGridView_videoListesi.Rows.Count.ToString();

                        // Bilgi Mesajı
                        statusLabel.ForeColor = Color.DarkGreen;
                        statusLabel.Text = "Videolarım Çekiliyor. Toplam: " + dataGridView_videoListesi.Rows.Count.ToString();


                    }
                }
                catch (Exception ex)
                {
                    // Bilgi Mesajı
                    statusLabel.ForeColor = Color.DarkRed;
                    statusLabel.Text = "Videolar çekilirken sorun oluştu. Hata: " + ex.Message;
                    listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Videolar çekilirken sorun oluştu. Hata: " + ex.Message);
                }

                videolarimiCek();
            }
            else
            {


                try
                {
                    // HtmlDocument sınıf tanımlama
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(kaynakKod);

                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//td[@class='videoListItem']");
                    foreach (var veri in XPath)
                    {
                        // veri
                        string video = veri.InnerHtml;

                        // video id
                        veriAyiklama(video, "/watch?v=", 9, "&amp;");
                        string video_id = ayiklananVeri;

                        // video adı
                        veriAyiklama(video, video_id, video_id.Length + 43, " </a>");
                        string video_adi = ayiklananVeri;

                        // Satır sayısını tanımlıyoruz
                        satirSayisi = dataGridView_videoListesi.Rows.Add();

                        // 1.satıra verileri yazdırıyoruz
                        dataGridView_videoListesi.Rows[satirSayisi].Cells[0].Value = satirSayisi + 1;
                        dataGridView_videoListesi.Rows[satirSayisi].Cells[1].Value = video_id;
                        dataGridView_videoListesi.Rows[satirSayisi].Cells[2].Value = video_adi;

                        // toplam video
                        label_topVideo.Text = dataGridView_videoListesi.Rows.Count.ToString();

                        // Bilgi Mesajı
                        statusLabel.ForeColor = Color.DarkGreen;
                        statusLabel.Text = "Videolarım Çekiliyor. Toplam: " + dataGridView_videoListesi.Rows.Count.ToString();
                    }
                }
                catch (Exception ex)
                {
                    // Bilgi Mesajı
                    statusLabel.ForeColor = Color.DarkRed;
                    statusLabel.Text = "Videolar çekilirken sorun oluştu. Hata: " + ex.Message;
                    listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Videolar çekilirken sorun oluştu. Hata: " + ex.Message);
                }



                // webbrowser durdur
                webBrowser2.Stop();
                //webBrowser2.Dispose();

                // thread durdur
                t_VideoCek.Abort();

                label_guncellemeZamani.Text = "Güncelleme Zamanı: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                // Bilgi Mesajı
                statusLabel.ForeColor = Color.DarkGreen;
                statusLabel.Text = "Tüm Videolar Çekildi. Toplam: " + dataGridView_videoListesi.Rows.Count.ToString();
                listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Tüm Videolar Çekildi. Toplam: " + dataGridView_videoListesi.Rows.Count.ToString());
                MessageBox.Show("Videoların Hepsi Çekildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        #endregion

        #region Datagrid Kolon Ekleme
        void listeKolonEkle()
        {
            // kolon temizle
            dataGridView_videoListesi.Columns.Clear();

            //Sütunları oluşturuyoruz.
            DataGridViewTextBoxColumn video_id = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn video_url = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn video_adi = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn video_paylasimZamani = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn video_durum = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn video_paylasimIcerigi = new DataGridViewTextBoxColumn();

            //Datagride alanlarımızı ekliyoruz.
            dataGridView_videoListesi.Columns.Add(video_id);
            dataGridView_videoListesi.Columns.Add(video_url);
            dataGridView_videoListesi.Columns.Add(video_adi);
            dataGridView_videoListesi.Columns.Add(video_paylasimZamani);
            dataGridView_videoListesi.Columns.Add(video_durum);
            dataGridView_videoListesi.Columns.Add(video_paylasimIcerigi);

            //Sütun başlıklarını ayarlıyoruz.
            video_id.HeaderText = "No";
            video_url.HeaderText = "Video Url";
            video_adi.HeaderText = "Video Adı";
            video_paylasimZamani.HeaderText = "Paylaşım Zamanı";
            video_durum.HeaderText = "Durum";
            video_paylasimIcerigi.HeaderText = "Paylaşım İçeriği";

            //Sütun genişliklerini ayarlıyoruz.
            video_id.Width = 25;
            video_url.Width = 140;
            video_adi.Width = 400;
            video_paylasimZamani.Width = 150;
            video_durum.Width = 100;
            video_paylasimIcerigi.Width = 100;

            // gizle bunu
            video_paylasimIcerigi.Visible = false;


        }
        #endregion

        #region Videolarımı Çek Butonu
        private void button_VideolarimiListele_Click(object sender, EventArgs e)
        {
            if (girisIslemi == true)
            {

                if (dataGridView_videoListesi.Rows.Count > 0)
                {
                    DialogResult soru = MessageBox.Show("Listede videolar mevcut. Tekrar Videolarım Listelensin mi?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (soru == DialogResult.Yes)
                    {

                        // kolon tanımla
                        listeKolonEkle();

                        // Bilgi Mesajı
                        statusLabel.ForeColor = Color.DarkGreen;
                        statusLabel.Text = "Videolarımı listele işlemi başladı.";
                        listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Videolarımı listele işlemi başladı.");
              
                        // toplam video sıfırla
                        label_topVideo.Text = "0";

                        // listeyi temizle
                        dataGridView_videoListesi.Rows.Clear();

                        // sayfayı sıfırla
                        sayfa = 0;

                        // thread başlatma
                        t_VideoCek = new Thread(delegate()
                        {
                            videolarimiCek();
                        });
                        t_VideoCek.Start();
                    }
                }
                else
                {
                    // kolon tanımla
                    listeKolonEkle();

                    // Bilgi Mesajı
                    statusLabel.ForeColor = Color.DarkGreen;
                    statusLabel.Text = "Videolarımı listele işlemi başladı.";
                    listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Videolarımı listele işlemi başladı.");
              
                    // toplam video sıfırla
                    label_topVideo.Text = "0";

                    // listeyi temizle
                    dataGridView_videoListesi.Rows.Clear();

                    // sayfayı sıfırla
                    sayfa = 0;

                    // thread başlatma
                    t_VideoCek = new Thread(delegate()
                    {
                        videolarimiCek();
                    });
                    t_VideoCek.Start();
                }
            }
            else
            {
                MessageBox.Show("Videoların listelenmesi için Youtube'a giriş yapmalısınız.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #endregion

        #region Seçili video bilgileri
        private void dataGridView_videoListesi_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // seçili
                groupBox_SeciliVideo.Text = "[" + dataGridView_videoListesi.Rows[e.RowIndex].Cells[0].Value.ToString() + "] Seçili Video Ayarları";
                textBox_VideoUrl.Text = "https://youtu.be/" + dataGridView_videoListesi.Rows[e.RowIndex].Cells[1].Value.ToString();
                textBox_videoAdi.Text = dataGridView_videoListesi.Rows[e.RowIndex].Cells[2].Value.ToString();

                // paylaşım durumu
                try
                {
                    textBox_PaylasimDurumu.Text = dataGridView_videoListesi.Rows[e.RowIndex].Cells[4].Value.ToString();
                }
                catch
                {
                    textBox_PaylasimDurumu.Text = "Paylaşılmadı";
                }

                // paylaşım içerği
                try
                {
                    string icerik = dataGridView_videoListesi.Rows[e.RowIndex].Cells[5].Value.ToString();
                    if (icerik.Length > 0)
                    {
                        richTextBox_VideoPaylasim.Text = icerik;
                    }
                    else
                    {
                        richTextBox_VideoPaylasim.Text = textBox_videoAdi.Text + "\n\n" + textBox_VideoUrl.Text + "\n\n" + "#hashtag ekle buraya :)";
                    }
                }
                catch
                {
                    richTextBox_VideoPaylasim.Text = textBox_videoAdi.Text + "\n\n" + textBox_VideoUrl.Text + "\n\n" + "#hashtag ekle buraya :)";
                }

                // zaman
                maskedTextBox_PaylasimZamani.Clear();
                try
                {
                    maskedTextBox_PaylasimZamani.Text = dataGridView_videoListesi.Rows[e.RowIndex].Cells[3].Value.ToString();

                }
                catch
                {
                    maskedTextBox_PaylasimZamani.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm");          
                }

                // Bilgi Mesajı
                statusLabel.ForeColor = Color.DarkBlue;
                statusLabel.Text = "[" + dataGridView_videoListesi.Rows[e.RowIndex].Cells[0].Value.ToString() + "] " + textBox_videoAdi.Text + " video düzenleniyor.";
                listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] [" + dataGridView_videoListesi.Rows[e.RowIndex].Cells[0].Value.ToString() + "] " + textBox_videoAdi.Text + " video düzenleniyor.");


            }
            catch (Exception ex)
            {
                MessageBox.Show("Video seçilemedi. Hata: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Karakter Limiti
        private void richTextBox_VideoPaylasim_KeyUp(object sender, KeyEventArgs e)
        {
            int limit = 140;
            int text_Limit = richTextBox_VideoPaylasim.TextLength;

            if (text_Limit > limit)
            {
                label_karakterSayisi.ForeColor = Color.Red;
                label_karakterSayisi.Text = richTextBox_VideoPaylasim.TextLength.ToString() + " karakter. " + limit.ToString() + " karakter limitini aştınız!";
            }
            else
            {
                label_karakterSayisi.ForeColor = Color.Blue;
                label_karakterSayisi.Text = richTextBox_VideoPaylasim.TextLength.ToString() + " karakter.";
            }
        }
        #endregion

        #region Listeyi Kaydet Butonu
        private void button_ListeyiKaydet_Click(object sender, EventArgs e)
        {
            // veri yoksa
            if (dataGridView_videoListesi.Rows.Count > 0)
            {
                #region kaydedilsin mi ?
                DialogResult soru = MessageBox.Show("Listede ki videolar Excel'e aktarılsın mı?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (soru == DialogResult.Yes)
                {
                    Excel.Application excel = new Excel.Application();
                    excel.Visible = true;
                    object Missing = Type.Missing;
                    Workbook workbook = excel.Workbooks.Add(Missing);
                    Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
                    int StartCol = 1;
                    int StartRow = 1;
                    for (int j = 0; j < dataGridView_videoListesi.Columns.Count; j++)
                    {
                        Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                        myRange.Value2 = dataGridView_videoListesi.Columns[j].HeaderText;
                    }
                    StartRow++;
                    for (int i = 0; i < dataGridView_videoListesi.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView_videoListesi.Columns.Count; j++)
                        {

                            Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                            myRange.Value2 = dataGridView_videoListesi[j, i].Value == null ? "" : dataGridView_videoListesi[j, i].Value;
                            myRange.Select();
                        }
                    }

                    // Bilgi Mesajı
                    statusLabel.ForeColor = Color.DarkGreen;
                    statusLabel.Text = "Listede ki videolar Excel'e aktarıldı.";
                    listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Listede ki videolar Excel'e aktarıldı.");


                }
                #endregion
            }
            else
            {
                MessageBox.Show("Liste de Excel'e aktarılacak video yok.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #endregion

        #region Liste Yükleme Fonksiyonu
        void videoListesiYukle()
        {
            try
            {
                // liste temizle önce
                dataGridView_videoListesi.Columns.Clear();
                dataGridView_videoListesi.Rows.Clear();

                // dosya yükleme penceresi
                OpenFileDialog yukle = new OpenFileDialog();
                yukle.Title = "Liste yükle";
                yukle.FileName = "";
                yukle.Filter = "Excel |*.xlsx";
                DialogResult ac = yukle.ShowDialog();

                // soru
                if (ac == DialogResult.OK)
                {
                    // bağlantılar
                    OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + yukle.FileName + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");

                    baglanti.Open();
                    OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sayfa1$]", baglanti);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    da.Fill(dt);
                    dataGridView_videoListesi.DataSource = dt.DefaultView;
                    baglanti.Close();
                }
            }
            catch
            {
                MessageBox.Show("Liste yüklemesinde sorun oluştu.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Video Listesi Yükleme Butonu
        private void button_ListeYukle_Click(object sender, EventArgs e)
        {
            // veri yoksa
            if (dataGridView_videoListesi.Rows.Count > 0)
            {
                DialogResult soru = MessageBox.Show("Liste de video var. Yinede Eklensin mi?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (soru == DialogResult.Yes)
                {
                    // video listesi yükle
                    videoListesiYukle();
                }
            }
            else
            {
                // video listesi yükle
                videoListesiYukle();
            }
        }
        #endregion

        #region Seçili Video Ayarlarını Kaydet Butonu
        private void button_SeciliVideoKaydet_Click(object sender, EventArgs e)
        {
            try
            {
                // video id al
                int videoID = dataGridView_videoListesi.CurrentCell.RowIndex;

                // durum
                if (textBox_PaylasimDurumu.Text.Length > 0)
                {
                    dataGridView_videoListesi.Rows[videoID].Cells[4].Value = textBox_PaylasimDurumu.Text;
                }
                else
                {
                    textBox_PaylasimDurumu.Text = "Paylaşılmadı";
                }

                // paylaşım içeriği
                dataGridView_videoListesi.Rows[videoID].Cells[5].Value = richTextBox_VideoPaylasim.Text;

                // paylaşım saat
                dataGridView_videoListesi.Rows[videoID].Cells[3].Value = maskedTextBox_PaylasimZamani.Text;


                // Bilgi Mesajı
                statusLabel.ForeColor = Color.DarkBlue;
                statusLabel.Text = "[" + dataGridView_videoListesi.Rows[videoID].Cells[0].Value.ToString() + "] " + textBox_videoAdi.Text + " seçili video güncellendi.";
                listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] [" + dataGridView_videoListesi.Rows[videoID].Cells[0].Value.ToString() + "] " + textBox_videoAdi.Text + " seçili video güncellendi.");
            }
            catch
            {
                MessageBox.Show("Kaydetmek için önce video seçmelisin.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #endregion

        #region Kontrol Aracı
        private void timer1_Tick(object sender, EventArgs e)
        {
            // güncel zaman
            label_zaman.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm").ToString();

            // Paylaşma Motoru Aktif mi ?
            if (paylasimMotoru == true)
            {
                try
                {
                    // listede video varmı
                    if (dataGridView_videoListesi.Rows.Count > 0)
                    {

                        // toplam video sayısı
                        int topVideo = dataGridView_videoListesi.Rows.Count;

                        // seçili video bilgileri
                        string videoAdi, videoPZamani, videoDurum, videoIcerik;

                        // for döngüsü
                        for (int i = 0; i < topVideo; i++)
                        {
                            // video adı
                            videoAdi = dataGridView_videoListesi.Rows[i].Cells[2].Value.ToString();

                            // paylasim zamanı
                            videoPZamani = dataGridView_videoListesi.Rows[i].Cells[3].Value.ToString();

                            // durum
                            videoDurum = dataGridView_videoListesi.Rows[i].Cells[4].Value.ToString();

                            // icerik
                            videoIcerik = dataGridView_videoListesi.Rows[i].Cells[5].Value.ToString();

                            // video paylaşılmadıysa paylaş
                            if (videoDurum == "Paylaşılmadı")
                            {
                                // paylaşma saati
                                string zaman = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
                                DateTime paylasimZamani = Convert.ToDateTime(videoPZamani);
                                DateTime simdikiZaman = Convert.ToDateTime(zaman);

                                // şimdiki zaman paylaşma zamanını geçmiş mi?
                                if (simdikiZaman > paylasimZamani)
                                {
                                    // twitter ve facebookta paylaş
                                    oto_Twitter(videoIcerik, videoAdi);


                                    // paylaşıldı olarak değiştir
                                    dataGridView_videoListesi.Rows[i].Cells[4].Value = "Paylaşıldı";

                                }
                            }
                        }
                    }
                    else
                    {
                        // nesne durumu
                        button_durdur.Enabled = false;
                        button_baslat.Enabled = true;

                        // paylaşmaya başla
                        paylasimMotoru = false;

                        // Bilgi Mesajı
                        statusLabel.ForeColor = Color.DarkBlue;
                        statusLabel.Text = "Sosyal Medya paylaşma motoru durdu. Liste de video bulunmuyor.";
                        listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Sosyal Medya paylaşma motoru durdu. Liste de video bulunmuyor.");
                        MessageBox.Show("Sosyal Medya paylaşma motoru durdu. Liste de video bulunmuyor.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    // nesne durumu
                    button_durdur.Enabled = false;
                    button_baslat.Enabled = true;

                    // paylaşmaya başla
                    paylasimMotoru = false;

                    // Bilgi Mesajı
                    statusLabel.ForeColor = Color.DarkRed;
                    statusLabel.Text = "Sosyal Medya paylaşma motoru durdu. Bir hata oluştu. HATA: " + ex.Message;
                    listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Sosyal Medya paylaşma motoru durdu. Bir hata oluştu.  HATA: " + ex.Message);
                    MessageBox.Show("Sosyal Medya paylaşma motoru durdu. Bir hata oluştu. HATA: " + ex.Message, "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                   
                }
            }
        }
        #endregion
        
        #region Başlat Butonu
        private void button_baslat_Click(object sender, EventArgs e)
        {
            // nesne durumu
            button_durdur.Enabled = true;
            button_baslat.Enabled = false;

            // paylaşmaya başla
            paylasimMotoru = true;

            // Bilgi Mesajı
            statusLabel.ForeColor = Color.DarkGreen;
            statusLabel.Text = "Sosyal Medya paylaşma motoru başladı.";
            listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Sosyal Medya paylaşma motoru başladı.");

        }
        #endregion

        #region Durdur Butonu
        private void button_durdur_Click(object sender, EventArgs e)
        {
            DialogResult soru = MessageBox.Show("Otomatik Paylaşımı durdurmak istiyor musunuzu?","Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if(soru == DialogResult.Yes)
            {
                // nesne durumu
                button_durdur.Enabled = false;
                button_baslat.Enabled = true;

                // paylaşmaya başla
                paylasimMotoru = false;

                // Bilgi Mesajı
                statusLabel.ForeColor = Color.DarkRed;
                statusLabel.Text = "Sosyal Medya paylaşma motoru durduruldu.";
                listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Sosyal Medya paylaşma motoru durduruldu.");  
               
            }
        }
        #endregion

        void oto_Twitter(string mesaj, string videoADI)
        {
            try
            {
                // api bilgileri
                var ConsumerKey = textBox_TW_consumerKey.Text;
                var ConsumerSecret = textBox_TW_consumerSecret.Text;
                var Token = textBox_TW_accessToken.Text;
                var TokenSecret = textBox_TW_accessTokenSecret.Text;

                //Twitter API servisine ConsumerKey ve ConsumerSecret bilgilerini girdik.
                var service = new TwitterService(ConsumerKey, ConsumerSecret);

                //Twitter API servisine Token ve TokenSecret bilgileri ile giriş yapıyoruz.
                service.AuthenticateWith(Token, TokenSecret);

                //Tweet gönder
                var result = service.SendTweet(new SendTweetOptions {
                    Status = mesaj //Tweet’in içeriğini giriyoruz.
                });


                // Bilgi Mesajı
                statusLabel.ForeColor = Color.DarkBlue;
                statusLabel.Text = videoADI + " video Twitter ve Facebook'ta paylaşıldı";
                listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] " + videoADI + " video Twitter ve Facebook'ta paylaşıldı");
                                
            }
            catch
            {
                // Bilgi Mesajı
                statusLabel.ForeColor = Color.DarkRed;
                statusLabel.Text = "Tweet atarken bir hata oluştu.";
                listBox_durum.Items.Insert(0, "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Tweet atarken bir hata oluştu.");
                MessageBox.Show("Tweet atarken bir hata oluştu.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                   
            }

        }

        #region Test Butonu
        private void button_TEST_Click(object sender, EventArgs e)
        {
            // özel mesaj gönderme DM

            // api bilgileri
            var ConsumerKey = textBox_TW_consumerKey.Text;
            var ConsumerSecret = textBox_TW_consumerSecret.Text;
            var Token = textBox_TW_accessToken.Text;
            var TokenSecret = textBox_TW_accessTokenSecret.Text;

            //Twitter API servisine ConsumerKey ve ConsumerSecret bilgilerini girdik.
            var service = new TwitterService(ConsumerKey, ConsumerSecret);

            //Twitter API servisine Token ve TokenSecret bilgileri ile giriş yapıyoruz.
            service.AuthenticateWith(Token, TokenSecret);


            var sonuc = service.SendDirectMessage(new SendDirectMessageOptions() {
                ScreenName = "vaftizcimusa", 
                Text = "deneme"
            });


        
        }
        #endregion



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // osman
            if (comboBox1.Text == "Osman Yavuz")
            {
                textBox_TW_consumerKey.Text = "yMPkviUIBdb0UG2l954Bceya5";
                textBox_TW_consumerSecret.Text = "JyoRfbr6Om1iuPiLv6088S5LYpOwTlhapKBviHqICnajxkoSN2";
                textBox_TW_accessToken.Text = "798634085333737472-SqktlNV6HJExkHOMZlh1zapIVD3My1h";
                textBox_TW_accessTokenSecret.Text = "NGyrhWRCpb1hlyBh5ocrEWTccB9549h4RVAqQbmt9l7Jg";     
            }
            else if (comboBox1.Text == "Vaftizci Musa")
            {
                textBox_TW_consumerKey.Text = "z6cdpco76bRuAJtUV7APrg3XW";
                textBox_TW_consumerSecret.Text = "YzzdZoVdpWmJXxtneLcSVKbwgUUSxz5qRPQmjH5W5qSYH9AZKt";
                textBox_TW_accessToken.Text = "803617960459833344-g22J3MgUdaX5dje6LpACB5sU39V8i54";
                textBox_TW_accessTokenSecret.Text = "g2O5iLphUh28ZuCt1aGgSWcFXkcTybTkF2LlAjWxEcvXR";
            }
            else
            {
                MessageBox.Show("Hesap seç lo");
            }
        }






    }
}
