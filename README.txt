E-ARŞİV UYGULAMASI - EXCEL EXPORT SÜRÜMÜ

Bu paket, PDF/ZIP faturalarını okuduktan sonra yüklediğiniz örnekle aynı
yapıda Excel dosyası üretir.

Excel sayfaları:
1. E-Fatura Özet
2. Kontrol
3. Özet

KURULUM

1. Çalışan Streamlit uygulamasını Ctrl + C ile durdurun.
2. Bu ZIP içindeki tüm dosya ve klasörleri mevcut e_arsiv_app klasörüne
   kopyalayın.
3. app.py için değiştirme uyarısı gelirse "Değiştir" seçin.
4. Terminalde sanal ortam aktifken şu komutu çalıştırın:

   python -m pip install -r requirements.txt

5. Uygulamayı başlatın:

   python -m streamlit run app.py

KULLANIM

1. PDF veya ZIP yükleyin.
2. "İşlemi Başlat" düğmesine basın.
3. İşlem bitince "Excel'i İndir" düğmesine basın.

ÖNEMLİ

templates/excel_sablonu.xlsx dosyasını silmeyin veya başka klasöre taşımayın.
Uygulama Excel biçimini bu şablondan alır.
