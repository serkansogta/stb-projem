Attribute VB_Name = "modKurulum"
'==============================================================================
' ASANSÖR PGD TAKİP SİSTEMİ - KURULUM VE BAŞLANGIÇ MODÜLÜ
' Modül : modKurulum
' Açıklama: İlk çalıştırmada tüm çalışma sayfalarını, başlıklarını,
'           Ribbon menüsünü, veri doğrulama listelerini ve AYARLAR
'           sayfasının varsayılan değerlerini oluşturur.
'
' KULLANIM: Excel'de Geliştirici > Makrolar > KurulumYap makrosunu çalıştırın.
'           Sadece bir kez çalıştırılması yeterlidir.
'==============================================================================
Option Explicit

'==============================================================================
' ANA KURULUM PROSEDÜRÜ
' Tüm kurulum adımlarını sırayla çalıştırır.
'==============================================================================
Public Sub KurulumYap()
    If Not Onayla("Asansör PGD Takip Sistemi kurulumu başlatılacak." & vbCrLf & _
                  "Bu işlem mevcut sayfa yapısını yeniden düzenleyecektir." & vbCrLf & vbCrLf & _
                  "Devam etmek istiyor musunuz?") Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 1. Sayfaları oluştur veya temizle
    SayfaOlustur SHT_GIRIS
    SayfaOlustur SHT_ASANSORLER
    SayfaOlustur SHT_PGD_KAYIT
    SayfaOlustur SHT_TAKVIM
    SayfaOlustur SHT_RAPORLAR
    SayfaOlustur SHT_AYARLAR
    SayfaOlustur SHT_SABLONLAR

    ' 2. Her sayfanın başlıklarını kur
    GirisSayfasiKur
    AsansorlerSayfasiKur
    PGDKayitlarSayfasiKur
    TakvimSayfasiKur
    RaporlarSayfasiKur
    AyarlarSayfasiKur
    SablonlarSayfasiKur

    ' 3. Veri doğrulama listelerini uygula
    VeriDogrulamaUygula

    ' 4. İlk özet güncelle
    GirisSayfasiGuncelle

    ' 5. Giriş sayfasına dön
    ThisWorkbook.Worksheets(SHT_GIRIS).Activate

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Kurulum tamamlandı!" & vbCrLf & vbCrLf & _
           "Başlamadan önce lütfen AYARLAR sayfasından" & vbCrLf & _
           "kurum bilgilerinizi doldurunuz.", vbInformation, UYGULAMA_ADI
End Sub

'==============================================================================
' SAYFA OLUŞTUR / TEMİZLE
'==============================================================================
Private Sub SayfaOlustur(ByVal sSayfaAdi As String)
    Dim ws As Worksheet
    Set ws = SayfaBul(sSayfaAdi)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sSayfaAdi
    Else
        ws.Cells.Clear
        ws.Cells.Interior.ColorIndex = xlNone
        ws.Cells.Font.ColorIndex = xlAutomatic
        ws.Cells.Font.Bold = False
        ws.Columns.ColumnWidth = 15
    End If
End Sub

'==============================================================================
' GİRİŞ SAYFASI KURULUMU - DASHBOARD
'==============================================================================
Private Sub GirisSayfasiKur()
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_GIRIS)
    ws.Tab.Color = RGB(0, 70, 127)

    ' Sayfa başlığı
    With ws.Range("A1:F1")
        .Merge
        .Value = UYGULAMA_ADI
        .Font.Bold = True
        .Font.Size = 18
        .Font.Color = RGB(0, 70, 127)
        .HorizontalAlignment = xlCenter
        .RowHeight = 45
    End With

    With ws.Range("A2:F2")
        .Merge
        .Value = AyarOku("KURUM_ADI")
        .Font.Size = 12
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With

    ' Özet kartları
    ws.Range("A4:B4").Merge
    ws.Range("A4:B4").Value = "ENVANTER DURUMU"
    BaslikStiliUygula ws.Range("A4:B4")

    Dim basliklar() As String
    basliklar = Split("Toplam Asansör|Aktif|Mühürlü|Devre Dışı|  |GECİKMİŞ PGD|Kritik (0-15 gün)|Yaklaşıyor (16-60 gün)|PGD Güncel (>60 gün)|  |Son Güncelleme", "|")
    Dim i As Long
    For i = 0 To UBound(basliklar)
        ws.Cells(5 + i, 1).Value = basliklar(i)
        ws.Cells(5 + i, 1).Font.Bold = True
        ws.Cells(5 + i, 2).Value = 0  ' Sayılar GirisSayfasiGuncelle ile doldurulur
        ws.Cells(5 + i, 1).RowHeight = 22
    Next i

    ' GECİKMİŞ satır başlığı rengi
    BaslikStiliUygula ws.Range("A9:B9")
    ws.Range("A9:B9").Interior.Color = RENK_TEHLIKE
    ws.Range("A9").Value = "PGD DURUM"

    ' Düğmeler (makro butonu olarak)
    With ws
        .Range("D4").Value = "HIZLI ERİŞİM"
        BaslikStiliUygula .Range("D4:F4")

        .Range("D5").Value = ">> Yeni Asansör Ekle"
        .Range("D6").Value = ">> PGD Kaydı Ekle"
        .Range("D7").Value = ">> Matbu Yazı Üret"
        .Range("D8").Value = ">> Asansör Ara"
        .Range("D9").Value = ">> Takvim Güncelle"
        .Range("D10").Value = ">> Durum Raporu"
        .Range("D11").Value = ">> Toplu PGD Bildirimi"
        .Range("D12").Value = ">> Asansörleri Yenile"

        Dim btn As Object
        Dim bBtnYok As Boolean
        bBtnYok = (ws.Shapes.Count = 0)
        If bBtnYok Then
            DugmeEkle ws, "D5", "AsansorEkle",        "Yeni Asansör Ekle"
            DugmeEkle ws, "D6", "PGDKayitEkle",       "PGD Kaydı Ekle"
            DugmeEkle ws, "D7", "MatbuYaziUret",      "Matbu Yazı Üret"
            DugmeEkle ws, "D8", "AsansorAra",         "Asansör Ara"
            DugmeEkle ws, "D9", "TakvimGuncelle",     "Takvim Güncelle"
            DugmeEkle ws, "D10", "DurumRaporuUret",   "Durum Raporu"
            DugmeEkle ws, "D11", "TopluPGDBildirimi", "Toplu PGD Bildirimi"
            DugmeEkle ws, "D12", "AsansorleriYenile", "Asansörleri Yenile"
        End If
    End With

    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 5
    ws.Columns("D:F").ColumnWidth = 22
End Sub

'==============================================================================
' ASANSÖRLER SAYFASI KURULUMU
'==============================================================================
Private Sub AsansorlerSayfasiKur()
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_ASANSORLER)
    ws.Tab.Color = RGB(0, 128, 0)

    Dim basliklar() As Variant
    basliklar = Array("ID", "Bina Adı", "Adres", "İlçe", "Malik/Yönetici", _
                      "Telefon", "E-posta", "Tescil No", "Tip", "Kapasite(kg)", _
                      "Durak", "Montaj Firması", "Servis Firması", "Son PGD Tarihi", _
                      "PGD Sonucu", "PGD Belge No", "PGD Kuruluşu", "Sonraki PGD", _
                      "Durum", "Notlar", "Kayıt Tarihi")

    Dim genislikler() As Integer
    genislikler = Array(6, 25, 30, 12, 20, 15, 20, 13, 13, 10, 7, 20, 20, 14, 13, 13, 22, 14, 10, 30, 16)

    Dim i As Integer
    For i = 0 To UBound(basliklar)
        ws.Cells(1, i + 1).Value = basliklar(i)
        ws.Columns(i + 1).ColumnWidth = genislikler(i)
    Next i
    BaslikStiliUygula ws.Range("A1:" & SutunHarfi(UBound(basliklar) + 1) & "1")

    ' Satırları dondur
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' AutoFilter uygula
    ws.Range("A1").AutoFilter
End Sub

'==============================================================================
' PGD KAYITLAR SAYFASI KURULUMU
'==============================================================================
Private Sub PGDKayitlarSayfasiKur()
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_PGD_KAYIT)
    ws.Tab.Color = RGB(0, 0, 200)

    Dim basliklar() As Variant
    basliklar = Array("Kayıt ID", "Asansör ID", "Bina Adı", "Tescil No", _
                      "Denetim Tarihi", "Denetim Kuruluşu", "Belge/Denetim No", _
                      "Sonuç", "Uygunsuzluk Sayısı", "Uygunsuzluk Notları", _
                      "Belge Tarihi", "Kayıt Tarihi")

    Dim genislikler() As Integer
    genislikler = Array(8, 9, 25, 13, 14, 25, 16, 14, 16, 35, 12, 16)

    Dim i As Integer
    For i = 0 To UBound(basliklar)
        ws.Cells(1, i + 1).Value = basliklar(i)
        ws.Columns(i + 1).ColumnWidth = genislikler(i)
    Next i
    BaslikStiliUygula ws.Range("A1:" & SutunHarfi(UBound(basliklar) + 1) & "1")

    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").AutoFilter
End Sub

'==============================================================================
' TAKVİM SAYFASI KURULUMU
'==============================================================================
Private Sub TakvimSayfasiKur()
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_TAKVIM)
    ws.Tab.Color = RGB(200, 100, 0)

    With ws.Range("A1:K1")
        .Rows(0).RowHeight = 15
    End With

    ' Açıklama satırı
    ws.Range("A1:K1").Merge
    ws.Range("A1").Value = "PGD TAKVİMİ - Kırmızı: Gecikmiş | Turuncu: Kritik | Sarı: Yaklaşıyor | Yeşil: Uygun"
    ws.Range("A1").Interior.Color = RGB(240, 240, 240)
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").HorizontalAlignment = xlCenter

    Dim basliklar() As Variant
    basliklar = Array("ID", "Bina Adı", "Adres", "İlçe", "Tescil No", _
                      "Malik/Yönetici", "Telefon", "Son PGD Tarihi", _
                      "Sonraki PGD", "Durum Metni", "Asansör Durumu")

    Dim genislikler() As Integer
    genislikler = Array(6, 25, 30, 12, 13, 20, 15, 14, 14, 24, 12)

    Dim i As Integer
    For i = 0 To UBound(basliklar)
        ws.Cells(2, i + 1).Value = basliklar(i)
        ws.Columns(i + 1).ColumnWidth = genislikler(i)
    Next i
    BaslikStiliUygula ws.Range("A2:" & SutunHarfi(UBound(basliklar) + 1) & "2")
End Sub

'==============================================================================
' RAPORLAR SAYFASI KURULUMU
'==============================================================================
Private Sub RaporlarSayfasiKur()
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_RAPORLAR)
    ws.Tab.Color = RGB(128, 0, 128)
    ws.Range("A1").Value = "Rapor üretmek için: GİRİŞ sayfasından 'Durum Raporu' düğmesine tıklayın."
    ws.Range("A1").Font.Italic = True
    ws.Columns("A:F").ColumnWidth = 20
End Sub

'==============================================================================
' AYARLAR SAYFASI KURULUMU
'==============================================================================
Private Sub AyarlarSayfasiKur()
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_AYARLAR)
    ws.Tab.Color = RGB(128, 128, 128)

    ws.Range("A1").Value = "ANAHTAR"
    ws.Range("B1").Value = "DEĞER"
    ws.Range("C1").Value = "AÇIKLAMA"
    BaslikStiliUygula ws.Range("A1:C1")

    Dim ayarlar() As Variant
    ayarlar = Array( _
        Array("KURUM_ADI",      "...... Belediyesi",               "Kurumun tam adı"), _
        Array("KURUM_BIRIM",    "Ruhsat ve Denetim Müdürlüğü",     "Birim/Müdürlük adı"), _
        Array("KURUM_BIRIM_KOD", "RDM",                             "Birim kısa kodu (yazı sayısında kullanılır)"), _
        Array("KURUM_ADRES",    "...... Mahallesi ...... Cad. No:XX ..... / ......", "Kurum adresi"), _
        Array("KURUM_TEL",      "0 (XXX) XXX XX XX",               "Kurum telefonu"), _
        Array("KURUM_EPOSTA",   "info@......bel.tr",                "Kurum e-posta adresi"), _
        Array("YETKILI_ADI",    "Ad Soyad",                         "İmzalayacak yetkili adı soyadı"), _
        Array("YETKILI_UNVAN",  "Müdür",                            "Yetkili unvanı"), _
        Array("KAYIT_KLASORU",  "",                                  "Üretilen yazıların kaydedileceği klasör (boş=Excel klasörü altında 'Yazılar')"), _
        Array("EVRAK_SAYAC",    "0",                                 "Otomatik evrak sayacı (elle değiştirmeyin)") _
    )

    Dim i As Integer
    For i = 0 To UBound(ayarlar)
        ws.Cells(i + 2, 1).Value = ayarlar(i)(0)
        ws.Cells(i + 2, 2).Value = ayarlar(i)(1)
        ws.Cells(i + 2, 3).Value = ayarlar(i)(2)
        ws.Cells(i + 2, 1).Font.Bold = True
        ws.Rows(i + 2).RowHeight = 20
    Next i

    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 50
    ws.Columns("C").ColumnWidth = 45

    MsgBox "Lütfen AYARLAR sayfasındaki 'KURUM_ADI', 'YETKILI_ADI' vb. bilgileri " & _
           "kendi kurumunuza göre güncelleyiniz.", vbInformation, UYGULAMA_ADI
End Sub

'==============================================================================
' ŞABLONLAR SAYFASI KURULUMU
' Resmi yazı şablonlarını <<DEĞIŞKEN>> yer tutucularıyla yükler.
'==============================================================================
Private Sub SablonlarSayfasiKur()
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_SABLONLAR)
    ws.Tab.Color = RGB(0, 128, 128)

    ws.Range("A1").Value = "ŞABLON_ADI"
    ws.Range("B1").Value = "İÇERİK (satır sonu için Alt+Enter kullanın)"
    ws.Range("C1").Value = "AÇIKLAMA"
    BaslikStiliUygula ws.Range("A1:C1")

    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 120
    ws.Columns("C").ColumnWidth = 30

    ' Şablonları yükle
    Dim satirlar() As Variant
    satirlar = Array( _
        Array("PGD_BILDIRIMI", SablonMetniPGDBildirimi(), "PGD Bildirimi şablonu"), _
        Array("IHTAR_YAZISI",  SablonMetniIhtar(),        "İhtar/Uyarı Yazısı şablonu"), _
        Array("MUHÜRLEME_KARARI", SablonMetniMuhürleme(), "Mühürleme Kararı şablonu"), _
        Array("UYGUNLUK_BELGESI", SablonMetniUygunluk(),  "Uygunluk Belgesi şablonu") _
    )

    Dim i As Integer
    For i = 0 To UBound(satirlar)
        ws.Cells(i + 2, 1).Value = satirlar(i)(0)
        ws.Cells(i + 2, 2).Value = satirlar(i)(1)
        ws.Cells(i + 2, 3).Value = satirlar(i)(2)
        ws.Cells(i + 2, 1).Font.Bold = True
        ws.Rows(i + 2).RowHeight = 60
        ws.Cells(i + 2, 2).WrapText = True
    Next i
End Sub

'==============================================================================
' VERİ DOĞRULAMA LİSTELERİ
'==============================================================================
Private Sub VeriDogrulamaUygula()
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_ASANSORLER)

    ' Asansör Tipi sütunu (I)
    With ws.Range("I2:I10000").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Operator:=1, _
             Formula1:=TIP_ELEKTRIKLI & "," & TIP_HIDROLIK & "," & TIP_ESCALATOR & "," & TIP_YURUMEYOLU
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' PGD Sonucu sütunu (O)
    With ws.Range("O2:O10000").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Operator:=1, _
             Formula1:=SONUC_UYGUN & "," & SONUC_UYGUNSUZ & "," & SONUC_KISMI
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' Durum sütunu (S)
    With ws.Range("S2:S10000").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Operator:=1, _
             Formula1:=DURUM_AKTIF & "," & DURUM_MUHURLU & "," & DURUM_DEVRE_DISI
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' PGD_KAYITLAR - Sonuç sütunu (H)
    Dim wsP As Worksheet
    Set wsP = SayfaBul(SHT_PGD_KAYIT)
    With wsP.Range("H2:H10000").Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Operator:=1, _
             Formula1:=SONUC_UYGUN & "," & SONUC_UYGUNSUZ & "," & SONUC_KISMI
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

'==============================================================================
' MAKRO DÜĞME EKLE (GİRİŞ SAYFASI İÇİN)
'==============================================================================
Private Sub DugmeEkle(ByVal ws As Worksheet, ByVal sHucre As String, _
                      ByVal sMakro As String, ByVal sYazi As String)
    On Error Resume Next
    Dim rng As Range
    Set rng = ws.Range(sHucre)
    Dim btn As Object
    Set btn = ws.Buttons.Add(rng.Left, rng.Top, rng.Width * 2 + 10, rng.Height)
    With btn
        .Caption = sYazi
        .OnAction = sMakro
        .Font.Size = 10
        .Font.Bold = False
    End With
    On Error GoTo 0
End Sub

'==============================================================================
' İŞ KİTABI OLAYLARI İÇİN YARDIMCI
' Bu kod ThisWorkbook modülüne taşınmalıdır.
'==============================================================================
Public Sub KitapAcildi_Cagri()
    ' GirisSayfasiGuncelle prosedürünü çağır
    GirisSayfasiGuncelle
End Sub

'==============================================================================
' ŞABLON METİNLERİ
' Tüm şablon metinleri bu bölümde tanımlanır.
' Değişkenler: <<DEĞİŞKEN_ADI>> formatında
'==============================================================================
Private Function SablonMetniPGDBildirimi() As String
    Dim s As String
    s = "<<KURUM_ADI>>" & Chr(10) & _
        "<<KURUM_BIRIM>>" & Chr(10) & _
        Chr(10) & _
        "Sayı   : <<EVRAK_SAYISI>>" & Chr(10) & _
        "Tarih  : <<TARIH>>" & Chr(10) & _
        "Konu   : Periyodik Güvenlik Denetimi (PGD) Yükümlülüğü Hakkında" & Chr(10) & _
        Chr(10) & _
        "<<MALIK_ADI>>" & Chr(10) & _
        "<<ADRES>>" & Chr(10) & _
        Chr(10) & _
        "İLGİ: 2014/33/EU Asansör Yönetmeliği ve ilgili mevzuat." & Chr(10) & _
        Chr(10) & _
        "<<ADRES>> adresinde bulunan <<BINA_ADI>> binaında tescil edilmiş <<TESCIL_NO>> " & _
        "sicil numaralı <<ASANSOR_TIP>> tipindeki asansörünüzün Periyodik Güvenlik " & _
        "Denetimi (PGD) yapılması yasal zorunluluktur." & Chr(10) & _
        Chr(10) & _
        "Kayıtlarımıza göre söz konusu asansörünüzün son PGD tarihi <<SON_PGD_TAR>> " & _
        "olup denetim süresi dolmuştur veya yaklaşmaktadır." & Chr(10) & _
        Chr(10) & _
        "6331 sayılı İş Sağlığı ve Güvenliği Kanunu ve Asansör Yönetmeliği " & _
        "çerçevesinde, <<SON_TAR>> tarihine kadar TÜRKAK onaylı A-Tipi Muayene " & _
        "Kuruluşuna PGD yaptırmanız ve sonuç belgelerini Müdürlüğümüze iletmeniz " & _
        "gerekmektedir." & Chr(10) & _
        Chr(10) & _
        "Belirtilen süre içinde PGD yaptırılmaması halinde asansörünüz mühürlenerek " & _
        "faaliyetten alınacak ve yasal işlem başlatılacaktır." & Chr(10) & _
        Chr(10) & _
        "Bilgilerinize saygıyla arz ederim." & Chr(10) & _
        Chr(10) & _
        "                                    " & Chr(10) & _
        "<<YETKILI_ADI>>" & Chr(10) & _
        "<<YETKILI_UNVAN>>" & Chr(10) & _
        "<<KURUM_BIRIM>>" & Chr(10) & _
        "<<KURUM_ADI>>"
    SablonMetniPGDBildirimi = s
End Function

Private Function SablonMetniIhtar() As String
    Dim s As String
    s = "<<KURUM_ADI>>" & Chr(10) & _
        "<<KURUM_BIRIM>>" & Chr(10) & _
        Chr(10) & _
        "Sayı   : <<EVRAK_SAYISI>>" & Chr(10) & _
        "Tarih  : <<TARIH>>" & Chr(10) & _
        "Konu   : PGD Yaptırılmaması Nedeniyle İHTAR" & Chr(10) & _
        Chr(10) & _
        "<<MALIK_ADI>>" & Chr(10) & _
        "<<ADRES>>" & Chr(10) & _
        Chr(10) & _
        "İHTAR" & Chr(10) & _
        Chr(10) & _
        "<<ADRES>> adresinde bulunan <<BINA_ADI>> binaında tescil edilmiş <<TESCIL_NO>> " & _
        "sicil numaralı asansörünüzün <<PGD_SON_TAR>> tarihinde sona eren Periyodik " & _
        "Güvenlik Denetimi (PGD) <<GECIKME_GUN>> hâlâ yaptırılmamıştır." & Chr(10) & _
        Chr(10) & _
        "Daha önce tarafınıza yapılan yazılı bildirime rağmen PGD yükümlülüğünü " & _
        "yerine getirmediğiniz anlaşılmaktadır. Bu ihmalden doğan sorumluluk " & _
        "tamamen size aittir." & Chr(10) & _
        Chr(10) & _
        "Bu yazının tarafınıza tebliğinden itibaren <<SURE_GUN>> (<<SURE_SON_TAR>> tarihine " & _
        "kadar) A-Tipi Muayene Kuruluşuna PGD yaptırarak belgeyi Müdürlüğümüze " & _
        "iletmeniz zorunludur." & Chr(10) & _
        Chr(10) & _
        "Süre içinde PGD yaptırılmaması halinde:" & Chr(10) & _
        "   - Asansörünüz mühürlenerek faaliyetten alınacaktır." & Chr(10) & _
        "   - <<CEZA_TUTARI>> tutarında idari para cezası uygulanacaktır." & Chr(10) & _
        "   - Asansörünüzü kullanan kişilerin uğrayacağı zararlardan hukuki sorumluluğunuz doğacaktır." & Chr(10) & _
        Chr(10) & _
        "İlgili mevzuat: 6331 sayılı Kanun Md.13, 7036 sayılı Kanun, Asansör Yönetmeliği." & Chr(10) & _
        Chr(10) & _
        "İHTAR OLUNUR." & Chr(10) & _
        Chr(10) & _
        "<<YETKILI_ADI>>" & Chr(10) & _
        "<<YETKILI_UNVAN>>" & Chr(10) & _
        "<<KURUM_BIRIM>>" & Chr(10) & _
        "<<KURUM_ADI>>"
    SablonMetniIhtar = s
End Function

Private Function SablonMetniMuhürleme() As String
    Dim s As String
    s = "<<KURUM_ADI>>" & Chr(10) & _
        "<<KURUM_BIRIM>>" & Chr(10) & _
        Chr(10) & _
        "MÜHÜRLEME / FAALİYETTEN ALIMA KARARI" & Chr(10) & _
        Chr(10) & _
        "Karar No  : <<KARAR_NO>>" & Chr(10) & _
        "Karar Tar.: <<TARIH>>" & Chr(10) & _
        Chr(10) & _
        "Aşağıda bilgileri verilen asansör, 2014/33/EU Asansör Yönetmeliği ve " & _
        "ilgili mevzuat hükümleri gereğince MÜHÜRLENEREK FAALİYETTEN ALINMIŞTIR." & Chr(10) & _
        Chr(10) & _
        "ASANSÖR BİLGİLERİ:" & Chr(10) & _
        "  Bina Adı     : <<BINA_ADI>>" & Chr(10) & _
        "  Adres        : <<ADRES>>" & Chr(10) & _
        "  Tescil No    : <<TESCIL_NO>>" & Chr(10) & _
        "  Malik        : <<MALIK_ADI>>" & Chr(10) & _
        "  Asansör Tipi : <<ASANSOR_TIP>>" & Chr(10) & _
        "  Kapasite     : <<KAPASITE>>" & Chr(10) & _
        "  Durak Sayısı : <<DURAK_SAYISI>>" & Chr(10) & _
        "  Son PGD Tar. : <<SON_PGD_TAR>>" & Chr(10) & _
        Chr(10) & _
        "MÜHÜRLEME GEREKÇESİ:" & Chr(10) & _
        "<<GEREKCE>>" & Chr(10) & _
        Chr(10) & _
        "ÖNEMLI UYARI: Mühürlü asansörün kullanılması yasaktır. Mühürü bozmak " & _
        "veya asansörü kullandırmak Türk Ceza Kanunu kapsamında suç teşkil etmekte " & _
        "olup yasal işlem başlatılacaktır." & Chr(10) & _
        Chr(10) & _
        "Asansörün yeniden faaliyete alınabilmesi için A-Tipi Muayene Kuruluşundan " & _
        "PGD uygunluk belgesi alınması ve Müdürlüğümüze başvurulması zorunludur." & Chr(10) & _
        Chr(10) & _
        "Bu karar tebligatın yapıldığı tarihte yürürlüğe girer." & Chr(10) & _
        Chr(10) & _
        "<<YETKILI_ADI>>" & Chr(10) & _
        "<<YETKILI_UNVAN>>" & Chr(10) & _
        "<<KURUM_BIRIM>>" & Chr(10) & _
        "<<KURUM_ADI>>"
    SablonMetniMuhürleme = s
End Function

Private Function SablonMetniUygunluk() As String
    Dim s As String
    s = "<<KURUM_ADI>>" & Chr(10) & _
        "<<KURUM_BIRIM>>" & Chr(10) & _
        Chr(10) & _
        "ASANSÖR PERİYODİK GÜVENLİK DENETİMİ UYGUNLUK BELGESİ" & Chr(10) & _
        Chr(10) & _
        "Belge No  : <<BELGE_NO>>" & Chr(10) & _
        "Düzenlenme: <<TARIH>>" & Chr(10) & _
        Chr(10) & _
        "Aşağıda bilgileri verilen asansörün Periyodik Güvenlik Denetimi (PGD) " & _
        "mevzuata uygun şekilde tamamlanmış olup UYGUN bulunduğu bildirilir." & Chr(10) & _
        Chr(10) & _
        "ASANSÖR BİLGİLERİ:" & Chr(10) & _
        "  Bina Adı         : <<BINA_ADI>>" & Chr(10) & _
        "  Adres            : <<ADRES>>" & Chr(10) & _
        "  Tescil No        : <<TESCIL_NO>>" & Chr(10) & _
        "  Malik            : <<MALIK_ADI>>" & Chr(10) & _
        "  Asansör Tipi     : <<ASANSOR_TIP>>" & Chr(10) & _
        "  Kapasite         : <<KAPASITE>>" & Chr(10) & _
        "  Durak Sayısı     : <<DURAK_SAYISI>>" & Chr(10) & _
        Chr(10) & _
        "PGD BİLGİLERİ:" & Chr(10) & _
        "  Denetim Tarihi   : <<PGD_TAR>>" & Chr(10) & _
        "  Denetim Kuruluşu : <<PGD_KURULUS>>" & Chr(10) & _
        "  Belge No         : <<PGD_BELGE_NO>>" & Chr(10) & _
        "  Geçerlilik Tarihi: <<SONRAKI_PGD>>" & Chr(10) & _
        Chr(10) & _
        "Asansörün belirtilen geçerlilik tarihi içinde güvenli olarak kullanılabilir. " & _
        "Geçerlilik süresinin dolmasından en az 60 (altmış) gün önce yeniden PGD " & _
        "yaptırılması yasal zorunluluktur." & Chr(10) & _
        Chr(10) & _
        "<<YETKILI_ADI>>" & Chr(10) & _
        "<<YETKILI_UNVAN>>" & Chr(10) & _
        "<<KURUM_BIRIM>>" & Chr(10) & _
        "<<KURUM_ADI>>"
    SablonMetniUygunluk = s
End Function
