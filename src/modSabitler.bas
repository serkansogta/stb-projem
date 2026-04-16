Attribute VB_Name = "modSabitler"
'==============================================================================
' ASANSÖR PGD TAKİP SİSTEMİ - SABİTLER VE GENEL TANIMLAR
' Modül : modSabitler
' Açıklama: Uygulama genelinde kullanılan sabitler, sayfa isimleri ve
'           sütun numaraları tanımlanır.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' SAYFA İSİMLERİ
'------------------------------------------------------------------------------
Public Const SHT_GIRIS       As String = "GİRİŞ"
Public Const SHT_ASANSORLER  As String = "ASANSÖRLER"
Public Const SHT_PGD_KAYIT   As String = "PGD_KAYITLAR"
Public Const SHT_TAKVIM      As String = "TAKVİM"
Public Const SHT_RAPORLAR    As String = "RAPORLAR"
Public Const SHT_AYARLAR     As String = "AYARLAR"
Public Const SHT_SABLONLAR   As String = "ŞABLONLAR"

'------------------------------------------------------------------------------
' ASANSÖRLER SAYFASI SÜTUN NUMARALARI
'------------------------------------------------------------------------------
Public Const ASN_ID             As Integer = 1   ' A - Kayıt ID (Otomatik)
Public Const ASN_BINA_ADI       As Integer = 2   ' B - Bina Adı
Public Const ASN_ADRES          As Integer = 3   ' C - Adres
Public Const ASN_ILCE           As Integer = 4   ' D - İlçe
Public Const ASN_MALIK_ADI      As Integer = 5   ' E - Malik/Yönetici Adı
Public Const ASN_MALIK_TEL      As Integer = 6   ' F - Malik Telefonu
Public Const ASN_MALIK_EMAIL    As Integer = 7   ' G - Malik E-posta
Public Const ASN_TESCIL_NO      As Integer = 8   ' H - Tescil / Sicil No
Public Const ASN_TIP            As Integer = 9   ' I - Asansör Tipi
Public Const ASN_KAPASITE       As Integer = 10  ' J - Kapasite (kg)
Public Const ASN_DURAK_SAYISI   As Integer = 11  ' K - Durak Sayısı
Public Const ASN_MONTAJ_FIRMA   As Integer = 12  ' L - Montaj Firması
Public Const ASN_SERVIS_FIRMA   As Integer = 13  ' M - Servis/Bakım Firması
Public Const ASN_SON_PGD_TAR    As Integer = 14  ' N - Son PGD Tarihi
Public Const ASN_SON_PGD_SONUC  As Integer = 15  ' O - Son PGD Sonucu
Public Const ASN_SON_PGD_BELGE  As Integer = 16  ' P - Son PGD Belge No
Public Const ASN_SON_PGD_KURULUS As Integer = 17 ' Q - Son PGD Denetim Kuruluşu
Public Const ASN_SON_PGD_TAR2   As Integer = 18  ' R - Sonraki PGD Tarihi (Otomatik)
Public Const ASN_DURUM          As Integer = 19  ' S - Durum
Public Const ASN_NOTLAR         As Integer = 20  ' T - Notlar
Public Const ASN_KAYIT_TAR      As Integer = 21  ' U - Kayıt Tarihi

'------------------------------------------------------------------------------
' PGD KAYITLAR SAYFASI SÜTUN NUMARALARI
'------------------------------------------------------------------------------
Public Const PGD_ID             As Integer = 1   ' A - Kayıt ID (Otomatik)
Public Const PGD_ASANSOR_ID     As Integer = 2   ' B - Asansör ID (Referans)
Public Const PGD_BINA_ADI       As Integer = 3   ' C - Bina Adı (Referans)
Public Const PGD_TESCIL_NO      As Integer = 4   ' D - Tescil No (Referans)
Public Const PGD_DENETIM_TAR    As Integer = 5   ' E - Denetim Tarihi
Public Const PGD_KURULUS        As Integer = 6   ' F - Denetim Kuruluşu
Public Const PGD_DENETIM_NO     As Integer = 7   ' G - Denetim / Belge No
Public Const PGD_SONUC          As Integer = 8   ' H - Denetim Sonucu
Public Const PGD_UYGUNSUZ_SAY   As Integer = 9   ' I - Uygunsuzluk Sayısı
Public Const PGD_UYGUNSUZ_NOT   As Integer = 10  ' J - Uygunsuzluk Notları
Public Const PGD_BELGE_TAR      As Integer = 11  ' K - Belge/Rapor Tarihi
Public Const PGD_KAYIT_TAR      As Integer = 12  ' L - Sisteme Giriş Tarihi

'------------------------------------------------------------------------------
' ASANSÖR DURUM DEĞERLERİ
'------------------------------------------------------------------------------
Public Const DURUM_AKTIF        As String = "Aktif"
Public Const DURUM_MUHURLU      As String = "Mühürlü"
Public Const DURUM_DEVRE_DISI   As String = "Devre Dışı"

'------------------------------------------------------------------------------
' PGD SONUÇ DEĞERLERİ
'------------------------------------------------------------------------------
Public Const SONUC_UYGUN        As String = "Uygun"
Public Const SONUC_UYGUNSUZ     As String = "Uygunsuz"
Public Const SONUC_KISMI        As String = "Kısmen Uygun"

'------------------------------------------------------------------------------
' ASANSÖR TİPLERİ
'------------------------------------------------------------------------------
Public Const TIP_ELEKTRIKLI     As String = "Elektrikli"
Public Const TIP_HIDROLIK       As String = "Hidrolik"
Public Const TIP_ESCALATOR      As String = "Yürüyen Merdiven"
Public Const TIP_YURUMEYOLU     As String = "Yürüyen Yol"

'------------------------------------------------------------------------------
' RENKLERİ (BGR formatında - Excel renk kodu)
'------------------------------------------------------------------------------
Public Const RENK_BASLIK        As Long = 1447446   ' Koyu Mavi (#161657 benzeri)
Public Const RENK_BASLIK_YAZ    As Long = 16777215  ' Beyaz
Public Const RENK_UYARI         As Long = 49407      ' Turuncu
Public Const RENK_TEHLIKE       As Long = 255        ' Kırmızı
Public Const RENK_UYGUN         As Long = 5287936    ' Yeşil
Public Const RENK_GERI          As Long = 13434828   ' Açık Gri
Public Const RENK_ZEBRA1        As Long = 16777215   ' Beyaz
Public Const RENK_ZEBRA2        As Long = 15921906   ' Çok açık gri

'------------------------------------------------------------------------------
' PGD SÜRE SABİTLERİ (GÜN)
'------------------------------------------------------------------------------
' Türk mevzuatına göre PGD süresi 12 aydır (365 gün)
Public Const PGD_SURE_GUN       As Integer = 365
' Uyarı eşiği: Kaç gün kala bildirim yapılsın?
Public Const PGD_UYARI_ESIGI    As Integer = 60
' Kritik eşik: Kaç gün kala kırmızıya dönsün?
Public Const PGD_KRITIK_ESIGI   As Integer = 15

'------------------------------------------------------------------------------
' METİN SABİTLERİ
'------------------------------------------------------------------------------
Public Const UYGULAMA_ADI      As String = "Asansör PGD Takip Sistemi"
Public Const UYGULAMA_VER      As String = "1.0"
