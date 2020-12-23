using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace ExcelClient
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
            string fullPath = "C:\\Users\\TUNC\\source\\repos\\ExcelClient\\ExcelClient\\ExampleData\\kvkk.xlsx";
            using (var stream = File.Open(fullPath, FileMode.Open, FileAccess.Read))
            {
                Reader reader = new Reader(stream);
                //reader.GetRow(0, "Envanter örneği");
                //reader.GetColumn(1, "Envanter örneği");
                //reader.GetTable("Envanter örneği");

            }
            */

            DataTableManager dataTableManager = new DataTableManager();
            // columns
            dataTableManager.DataTable.Columns.Add(" ",typeof(int));
            dataTableManager.DataTable.Columns.Add("Departman", typeof(string));
            dataTableManager.DataTable.Columns.Add("Faaliyet", typeof(string));
            dataTableManager.DataTable.Columns.Add("Veri Kategorisi", typeof(string));
            dataTableManager.DataTable.Columns.Add("Kişisel Veri", typeof(string));
            dataTableManager.DataTable.Columns.Add("Özel Nitelikli Kişisel Veri", typeof(string));
            dataTableManager.DataTable.Columns.Add("İşleme Amacı", typeof(string));
            dataTableManager.DataTable.Columns.Add("Veri Konusu Kişi Grubu", typeof(string));
            dataTableManager.DataTable.Columns.Add("Hukuki Sebebi", typeof(string));
            dataTableManager.DataTable.Columns.Add("Saklama Süresi", typeof(string));
            dataTableManager.DataTable.Columns.Add("Alıcı / Alıcı Grupları", typeof(string));
            dataTableManager.DataTable.Columns.Add("Yabancı Ülkelere Aktarılan Veriler", typeof(string));
            dataTableManager.DataTable.Columns.Add("İdari Tedbirler", typeof(string));
            dataTableManager.DataTable.Columns.Add("Teknik Tedbirler", typeof(string));

            dataTableManager.DataTable.Rows.Add(1,"İnsan Kaynakları", "Çalışan Özlük " + Environment.NewLine+ "Dosyası Oluşturma" ,"Kimlik", "Ad,Soyad", " ", "Çalışanlar İçin İş Akdi ve Mevzuat Kaynaklı Yükümlülüklerin Yerine Getirilmesi", "Çalışanlar", "Sözleşme İmzalanması", "İşten Ayrılmasından İtibareen 10 Yıl", "SGK Ve Diğer Yetkili Kurum ve Kuruluşlar", "Yurt Dışına Aktarım Yapılmıyor", "Çalışanların Niteliği Ve Teknik Bilgi/Becerisinin Geliştirilmesi, Kişisel Verilerin Hukuka Aykırı İşlenmenin Önlenmesi, " + Environment.NewLine +" Kişisel Verilere Hukuka Aykırı Erişilmesinin Önlenmesi, Kişisel Verilerin Muhafazasının Sağlanması, İletişim Teknikleri Ve İlgili Mevzuatlar Hakkında Eğitimler Verilmekte; Çalışanlara Gizlilik Sözleşmeleri İmzalatılmakta; Güvenlik Politika Ve Prosedürlerine Uymayan Çalışanlara Yönelik Uygulanacak Disiplin Prosedürü Uygulanmakta, İlgili Kişileri Aydınlatma Yükümlülüğü Yerine Getirilmekte, Kurum İçi Periyodik Ve Rastgele Denetimler Yapılmakta Ve Çalışanlara Yönelik Bilgi Güvenliği Eğitimleri Verilmektedir.", "Kurumun Bilişim Sistemleri Teçhizatı, Yazılım Ve Verilerin Fiziksel Güvenliği İçin Gerekli Önlemler Alınmakta, Hukuka Aykırı İşlemeyi Önlemeye Yönelik Riskler Belirlenmekte, Bu Risklere Uygun Teknik Tedbirler Alınmakta, Erişim Yetki Ve Rol Dağılımları İçin Prosedürler Oluşturulmakta Ve Uygulanmakta, Yetki Matrisi Uygulanmakta, " + Environment.NewLine + " Erişimler Kayıt Altına Alınarak Uygunsuz Erişimler Kontrol Altında Tutulmakta, Saklama Ve İmha Politikasına Uygun İmha Süreçleri Tanımlanmakta Ve Uygulanmakta, Hukuka Aykırı İşleme Tespiti Halinde İlgili Kişiye Ve Kurula Bildirmek İçin Bir Sistem Ve Altyapı Oluşturulmakta, Güvenlik Açıkları Takip Edilerek Uygun Güvenlik Yamaları Yüklenmekte, Bilgi Sistemleri Güncel Halde Tutulmakta, Kişisel Verilerin İşlendiği Elektronik Ortamlarda Güçlü Parolalar Kullanılmakta Ve Güvenli Kayıt Tutma (Loglama) Sistemleri Kullanılmakta, Kişisel Verilerin Güvenli Olarak Saklanmasını Sağlayan Yedekleme Programları Kullanılmaktadır.");
            

            
            
            Writer writer = new Writer(dataTableManager.DataTable, "VERİ ENVANTERİ","KİŞİSEL VERİ İŞLEME ENVANTERİ", @"C:\Book1.xlsx");
            writer.CloseExcel();
            
            //Console.ReadKey();
        }
    }
}
