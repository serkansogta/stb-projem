Attribute VB_Name = "modEnvanter"
'==============================================================================
' ASANSÖR PGD TAKİP SİSTEMİ - ENVANTER YÖNETİMİ
' Modül : modEnvanter
' Açıklama: Asansör envanterinin ekleme, düzenleme, silme ve listeleme
'           işlemlerini yönetir.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' YENİ ASANSÖR EKLE
' Kullanıcıdan form üzerinden veri alarak ASANSÖRLER sayfasına yazar.
' Döndürür: Yeni kayıt ID'si (başarısızsa 0)
'------------------------------------------------------------------------------
Public Function AsansorEkle() As Long
    AsansorEkle = 0

    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_ASANSORLER)
    If ws Is Nothing Then
        MsgBox "ASANSÖRLER sayfası bulunamadı!", vbCritical, UYGULAMA_ADI
        Exit Function
    End If

    ' --- Kullanıcıdan veri al ---
    Dim sBinaAdi As String
    Dim sAdres As String
    Dim sIlce As String
    Dim sMalikAdi As String
    Dim sMalikTel As String
    Dim sMalikEmail As String
    Dim sTescilNo As String
    Dim sTip As String
    Dim sKapasite As String
    Dim sDurakSayisi As String
    Dim sMontajFirma As String
    Dim sServisFirma As String
    Dim sSonPGDTar As String
    Dim sSonPGDSonuc As String
    Dim sSonPGDBelge As String
    Dim sSonPGDKurulus As String
    Dim sDurum As String
    Dim sNotlar As String

    sBinaAdi = Trim(InputBox("Bina Adı:", UYGULAMA_ADI & " - Yeni Asansör"))
    If sBinaAdi = "" Then Exit Function

    sAdres = Trim(InputBox("Açık Adres:", UYGULAMA_ADI & " - Yeni Asansör", ""))
    sIlce = Trim(InputBox("İlçe:", UYGULAMA_ADI & " - Yeni Asansör", ""))
    sMalikAdi = Trim(InputBox("Malik / Yönetici Adı Soyadı:", UYGULAMA_ADI & " - Yeni Asansör", ""))
    sMalikTel = Trim(InputBox("Malik Telefonu:", UYGULAMA_ADI & " - Yeni Asansör", ""))
    sMalikEmail = Trim(InputBox("Malik E-posta (boş bırakılabilir):", UYGULAMA_ADI & " - Yeni Asansör", ""))
    sTescilNo = Trim(InputBox("Tescil / Sicil No:", UYGULAMA_ADI & " - Yeni Asansör", ""))

    ' Asansör tipi seçimi
    sTip = AsansorTipSec()
    If sTip = "" Then Exit Function

    sKapasite = Trim(InputBox("Kapasite (kg):", UYGULAMA_ADI & " - Yeni Asansör", ""))
    sDurakSayisi = Trim(InputBox("Durak Sayısı:", UYGULAMA_ADI & " - Yeni Asansör", ""))
    sMontajFirma = Trim(InputBox("Montaj Firması:", UYGULAMA_ADI & " - Yeni Asansör", ""))
    sServisFirma = Trim(InputBox("Servis/Bakım Firması:", UYGULAMA_ADI & " - Yeni Asansör", ""))

    sSonPGDTar = Trim(InputBox("Son PGD Tarihi (GG.AA.YYYY, boş bırakılabilir):", _
                               UYGULAMA_ADI & " - Yeni Asansör", ""))
    If sSonPGDTar <> "" Then
        If Not IsDate(sSonPGDTar) Then
            MsgBox "Geçersiz tarih formatı! Kayıt iptal edildi.", vbExclamation, UYGULAMA_ADI
            Exit Function
        End If
        sSonPGDSonuc = PGDSonucSec()
        sSonPGDBelge = Trim(InputBox("Son PGD Belge No:", UYGULAMA_ADI & " - Yeni Asansör", ""))
        sSonPGDKurulus = Trim(InputBox("Son PGD Denetim Kuruluşu:", UYGULAMA_ADI & " - Yeni Asansör", ""))
    End If

    sDurum = DurumSec()
    If sDurum = "" Then sDurum = DURUM_AKTIF
    sNotlar = Trim(InputBox("Notlar (isteğe bağlı):", UYGULAMA_ADI & " - Yeni Asansör", ""))

    ' --- Sayfaya yaz ---
    Dim nID As Long
    Dim nSatir As Long
    nID = YeniID(ws)
    nSatir = SonSatir(ws, 1) + 1

    SayfaKorumaKapat ws

    With ws
        .Cells(nSatir, ASN_ID).Value = nID
        .Cells(nSatir, ASN_BINA_ADI).Value = sBinaAdi
        .Cells(nSatir, ASN_ADRES).Value = sAdres
        .Cells(nSatir, ASN_ILCE).Value = sIlce
        .Cells(nSatir, ASN_MALIK_ADI).Value = sMalikAdi
        .Cells(nSatir, ASN_MALIK_TEL).Value = sMalikTel
        .Cells(nSatir, ASN_MALIK_EMAIL).Value = sMalikEmail
        .Cells(nSatir, ASN_TESCIL_NO).Value = sTescilNo
        .Cells(nSatir, ASN_TIP).Value = sTip
        If IsNumeric(sKapasite) Then .Cells(nSatir, ASN_KAPASITE).Value = CDbl(sKapasite)
        If IsNumeric(sDurakSayisi) Then .Cells(nSatir, ASN_DURAK_SAYISI).Value = CInt(sDurakSayisi)
        .Cells(nSatir, ASN_MONTAJ_FIRMA).Value = sMontajFirma
        .Cells(nSatir, ASN_SERVIS_FIRMA).Value = sServisFirma

        If IsDate(sSonPGDTar) And sSonPGDTar <> "" Then
            .Cells(nSatir, ASN_SON_PGD_TAR).Value = CDate(sSonPGDTar)
            .Cells(nSatir, ASN_SON_PGD_TAR).NumberFormat = "DD.MM.YYYY"
            .Cells(nSatir, ASN_SON_PGD_SONUC).Value = sSonPGDSonuc
            .Cells(nSatir, ASN_SON_PGD_BELGE).Value = sSonPGDBelge
            .Cells(nSatir, ASN_SON_PGD_KURULUS).Value = sSonPGDKurulus
            ' Sonraki PGD tarihi = Son PGD + 365 gün
            Dim dSonraki As Date
            dSonraki = CDate(sSonPGDTar) + PGD_SURE_GUN
            .Cells(nSatir, ASN_SON_PGD_TAR2).Value = dSonraki
            .Cells(nSatir, ASN_SON_PGD_TAR2).NumberFormat = "DD.MM.YYYY"
        End If

        .Cells(nSatir, ASN_DURUM).Value = sDurum
        .Cells(nSatir, ASN_NOTLAR).Value = sNotlar
        .Cells(nSatir, ASN_KAYIT_TAR).Value = Now
        .Cells(nSatir, ASN_KAYIT_TAR).NumberFormat = "DD.MM.YYYY HH:MM"
    End With

    ' Satır renklendirmesi
    AsansorSatiriBoyа ws, nSatir

    SayfaKorumaAc ws

    MsgBox "Asansör başarıyla eklendi!" & vbCrLf & _
           "Kayıt ID : " & nID & vbCrLf & _
           "Bina     : " & sBinaAdi, vbInformation, UYGULAMA_ADI

    AsansorEkle = nID
End Function

'------------------------------------------------------------------------------
' ASANSÖR DÜZENLE
' ID'ye göre mevcut asansör kaydını günceller.
'------------------------------------------------------------------------------
Public Sub AsansorDuzenle(Optional ByVal nID As Long = 0)
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_ASANSORLER)
    If ws Is Nothing Then Exit Sub

    If nID = 0 Then
        Dim sID As String
        sID = Trim(InputBox("Düzenlenecek Asansörün Kayıt ID'sini girin:", UYGULAMA_ADI))
        If sID = "" Or Not IsNumeric(sID) Then Exit Sub
        nID = CLng(sID)
    End If

    Dim nSatir As Long
    nSatir = AsansorSatirBul(ws, nID)
    If nSatir = 0 Then
        MsgBox "ID=" & nID & " olan asansör kaydı bulunamadı!", vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If

    ' Mevcut değerleri al ve güncelle
    SayfaKorumaKapat ws
    With ws
        Dim sYeni As String

        sYeni = Trim(InputBox("Bina Adı:", UYGULAMA_ADI & " - Düzenle", .Cells(nSatir, ASN_BINA_ADI).Value))
        If sYeni <> "" Then .Cells(nSatir, ASN_BINA_ADI).Value = sYeni

        sYeni = Trim(InputBox("Adres:", UYGULAMA_ADI & " - Düzenle", .Cells(nSatir, ASN_ADRES).Value))
        .Cells(nSatir, ASN_ADRES).Value = sYeni

        sYeni = Trim(InputBox("İlçe:", UYGULAMA_ADI & " - Düzenle", .Cells(nSatir, ASN_ILCE).Value))
        .Cells(nSatir, ASN_ILCE).Value = sYeni

        sYeni = Trim(InputBox("Malik/Yönetici:", UYGULAMA_ADI & " - Düzenle", .Cells(nSatir, ASN_MALIK_ADI).Value))
        .Cells(nSatir, ASN_MALIK_ADI).Value = sYeni

        sYeni = Trim(InputBox("Malik Telefonu:", UYGULAMA_ADI & " - Düzenle", .Cells(nSatir, ASN_MALIK_TEL).Value))
        .Cells(nSatir, ASN_MALIK_TEL).Value = sYeni

        sYeni = Trim(InputBox("Tescil No:", UYGULAMA_ADI & " - Düzenle", .Cells(nSatir, ASN_TESCIL_NO).Value))
        .Cells(nSatir, ASN_TESCIL_NO).Value = sYeni

        sYeni = Trim(InputBox("Servis Firması:", UYGULAMA_ADI & " - Düzenle", .Cells(nSatir, ASN_SERVIS_FIRMA).Value))
        .Cells(nSatir, ASN_SERVIS_FIRMA).Value = sYeni

        sYeni = Trim(InputBox("Durum (Aktif / Mühürlü / Devre Dışı):", _
                              UYGULAMA_ADI & " - Düzenle", .Cells(nSatir, ASN_DURUM).Value))
        .Cells(nSatir, ASN_DURUM).Value = sYeni

        sYeni = Trim(InputBox("Notlar:", UYGULAMA_ADI & " - Düzenle", .Cells(nSatir, ASN_NOTLAR).Value))
        .Cells(nSatir, ASN_NOTLAR).Value = sYeni
    End With

    AsansorSatiriBoyа ws, nSatir
    SayfaKorumaAc ws

    MsgBox "Kayıt güncellendi. (ID=" & nID & ")", vbInformation, UYGULAMA_ADI
End Sub

'------------------------------------------------------------------------------
' ASANSÖR SİL
' ID'ye göre asansör kaydını siler (onay alınarak).
'------------------------------------------------------------------------------
Public Sub AsansorSil(Optional ByVal nID As Long = 0)
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_ASANSORLER)
    If ws Is Nothing Then Exit Sub

    If nID = 0 Then
        Dim sID As String
        sID = Trim(InputBox("Silinecek Asansörün Kayıt ID'sini girin:", UYGULAMA_ADI))
        If sID = "" Or Not IsNumeric(sID) Then Exit Sub
        nID = CLng(sID)
    End If

    Dim nSatir As Long
    nSatir = AsansorSatirBul(ws, nID)
    If nSatir = 0 Then
        MsgBox "ID=" & nID & " olan asansör bulunamadı!", vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If

    Dim sBilgi As String
    sBilgi = "Bina : " & ws.Cells(nSatir, ASN_BINA_ADI).Value & vbCrLf & _
             "Tescil: " & ws.Cells(nSatir, ASN_TESCIL_NO).Value

    If Not Onayla("Aşağıdaki asansör kaydını SİLMEK istediğinizden emin misiniz?" & _
                  vbCrLf & vbCrLf & sBilgi & vbCrLf & vbCrLf & _
                  "Bu işlem GERİ ALINAMAZ!") Then Exit Sub

    SayfaKorumaKapat ws
    ws.Rows(nSatir).Delete
    SayfaKorumaAc ws

    MsgBox "Asansör kaydı silindi.", vbInformation, UYGULAMA_ADI
End Sub

'------------------------------------------------------------------------------
' ASANSÖR ARA
' Belirtilen kriterlere göre ASANSÖRLER sayfasını filtreler.
'------------------------------------------------------------------------------
Public Sub AsansorAra()
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_ASANSORLER)
    If ws Is Nothing Then Exit Sub

    Dim sArama As String
    sArama = Trim(InputBox("Arama metni girin (bina adı, adres, tescil no veya malik adı):", _
                           UYGULAMA_ADI & " - Asansör Ara", ""))
    If sArama = "" Then
        ' Filtreyi kaldır
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        MsgBox "Filtre kaldırıldı.", vbInformation, UYGULAMA_ADI
        Exit Sub
    End If

    ' AutoFilter ile tüm metin sütunlarında ara (bina adı, adres, tescil no)
    If Not ws.AutoFilterMode Then ws.Rows(1).AutoFilter

    ' Bina adında ara (sütun B)
    ws.AutoFilter.Filters.Item(ASN_BINA_ADI).On = False
    ws.Range("A1").AutoFilter Field:=ASN_BINA_ADI, Criteria1:="*" & sArama & "*"

    MsgBox "'" & sArama & "' için filtreleme uygulandı." & vbCrLf & _
           "Filtreyi kaldırmak için tekrar 'Asansör Ara' menüsünü kullanın.", _
           vbInformation, UYGULAMA_ADI
End Sub

'------------------------------------------------------------------------------
' ASANSÖR SATIR BUL (iç kullanım)
' ID'ye göre ASANSÖRLER sayfasındaki satır numarasını döndürür.
' Bulunamazsa 0 döner.
'------------------------------------------------------------------------------
Private Function AsansorSatirBul(ByVal ws As Worksheet, ByVal nID As Long) As Long
    Dim nSon As Long
    Dim i As Long
    nSon = SonSatir(ws, 1)
    For i = 2 To nSon
        If ws.Cells(i, ASN_ID).Value = nID Then
            AsansorSatirBul = i
            Exit Function
        End If
    Next i
    AsansorSatirBul = 0
End Function

'------------------------------------------------------------------------------
' ASANSÖR SATIRI RENKLENDİR (iç kullanım)
' Durum ve PGD tarihine göre satır arka plan rengini günceller.
'------------------------------------------------------------------------------
Private Sub AsansorSatiriBoyа(ByVal ws As Worksheet, ByVal nSatir As Long)
    Dim sAlt As Long
    sAlt = IIf(nSatir Mod 2 = 0, RENK_ZEBRA2, RENK_ZEBRA1)
    ws.Rows(nSatir).Interior.Color = sAlt

    ' Durum sütununu renklendir
    Dim sDurum As String
    sDurum = ws.Cells(nSatir, ASN_DURUM).Value
    Select Case sDurum
        Case DURUM_MUHURLU
            ws.Cells(nSatir, ASN_DURUM).Interior.Color = RENK_TEHLIKE
            ws.Cells(nSatir, ASN_DURUM).Font.Color = RENK_BASLIK_YAZ
        Case DURUM_DEVRE_DISI
            ws.Cells(nSatir, ASN_DURUM).Interior.Color = RENK_GERI
            ws.Cells(nSatir, ASN_DURUM).Font.Color = RGB(80, 80, 80)
        Case Else ' Aktif
            ws.Cells(nSatir, ASN_DURUM).Interior.Color = RENK_UYGUN
            ws.Cells(nSatir, ASN_DURUM).Font.Color = RENK_BASLIK_YAZ
    End Select

    ' Sonraki PGD tarih sütununu renklendir
    Dim dSonraki As Variant
    dSonraki = ws.Cells(nSatir, ASN_SON_PGD_TAR2).Value
    ws.Cells(nSatir, ASN_SON_PGD_TAR2).Interior.Color = PGDDurumRengi(dSonraki)
    If PGDDurumRengi(dSonraki) <> RENK_ZEBRA1 And PGDDurumRengi(dSonraki) <> RENK_ZEBRA2 Then
        ws.Cells(nSatir, ASN_SON_PGD_TAR2).Font.Color = RENK_BASLIK_YAZ
    End If
End Sub

'------------------------------------------------------------------------------
' TÜM ASANSÖR SATIRLARINI YENİDEN RENKLENDIR
' Sayfa açıldığında veya toplu güncelleme sonrasında çağrılır.
'------------------------------------------------------------------------------
Public Sub AsansorleriYenile()
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_ASANSORLER)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    SayfaKorumaKapat ws

    Dim nSon As Long
    Dim i As Long
    nSon = SonSatir(ws, 1)
    For i = 2 To nSon
        If ws.Cells(i, ASN_ID).Value <> "" Then
            ' Sonraki PGD tarihini yeniden hesapla
            Dim dSon As Variant
            dSon = ws.Cells(i, ASN_SON_PGD_TAR).Value
            If IsDate(dSon) And Not IsEmpty(dSon) Then
                ws.Cells(i, ASN_SON_PGD_TAR2).Value = CDate(dSon) + PGD_SURE_GUN
                ws.Cells(i, ASN_SON_PGD_TAR2).NumberFormat = "DD.MM.YYYY"
            End If
            AsansorSatiriBoyа ws, i
        End If
    Next i

    SayfaKorumaAc ws
    Application.ScreenUpdating = True
    MsgBox "Asansör listesi yenilendi. Toplam " & (nSon - 1) & " kayıt.", _
           vbInformation, UYGULAMA_ADI
End Sub

'------------------------------------------------------------------------------
' SEÇIM YARDIMCILARI (iç kullanım)
'------------------------------------------------------------------------------
Private Function AsansorTipSec() As String
    Dim sSecenekler As String
    sSecenekler = "1 - " & TIP_ELEKTRIKLI & vbCrLf & _
                  "2 - " & TIP_HIDROLIK & vbCrLf & _
                  "3 - " & TIP_ESCALATOR & vbCrLf & _
                  "4 - " & TIP_YURUMEYOLU
    Dim sSec As String
    sSec = Trim(InputBox("Asansör tipini seçin:" & vbCrLf & vbCrLf & sSecenekler, _
                         UYGULAMA_ADI & " - Asansör Tipi", "1"))
    Select Case sSec
        Case "1": AsansorTipSec = TIP_ELEKTRIKLI
        Case "2": AsansorTipSec = TIP_HIDROLIK
        Case "3": AsansorTipSec = TIP_ESCALATOR
        Case "4": AsansorTipSec = TIP_YURUMEYOLU
        Case Else: AsansorTipSec = TIP_ELEKTRIKLI
    End Select
End Function

Private Function PGDSonucSec() As String
    Dim sSecenekler As String
    sSecenekler = "1 - " & SONUC_UYGUN & vbCrLf & _
                  "2 - " & SONUC_UYGUNSUZ & vbCrLf & _
                  "3 - " & SONUC_KISMI
    Dim sSec As String
    sSec = Trim(InputBox("PGD sonucunu seçin:" & vbCrLf & vbCrLf & sSecenekler, _
                         UYGULAMA_ADI & " - PGD Sonucu", "1"))
    Select Case sSec
        Case "1": PGDSonucSec = SONUC_UYGUN
        Case "2": PGDSonucSec = SONUC_UYGUNSUZ
        Case "3": PGDSonucSec = SONUC_KISMI
        Case Else: PGDSonucSec = SONUC_UYGUN
    End Select
End Function

Private Function DurumSec() As String
    Dim sSecenekler As String
    sSecenekler = "1 - " & DURUM_AKTIF & vbCrLf & _
                  "2 - " & DURUM_MUHURLU & vbCrLf & _
                  "3 - " & DURUM_DEVRE_DISI
    Dim sSec As String
    sSec = Trim(InputBox("Durum seçin:" & vbCrLf & vbCrLf & sSecenekler, _
                         UYGULAMA_ADI & " - Durum", "1"))
    Select Case sSec
        Case "1": DurumSec = DURUM_AKTIF
        Case "2": DurumSec = DURUM_MUHURLU
        Case "3": DurumSec = DURUM_DEVRE_DISI
        Case Else: DurumSec = DURUM_AKTIF
    End Select
End Function

'------------------------------------------------------------------------------
' ASANSÖR BİLGİSİ GETİR
' ID'ye göre asansör verilerini bir sözlük (Collection) olarak döndürür.
'------------------------------------------------------------------------------
Public Function AsansorGetir(ByVal nID As Long) As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_ASANSORLER)
    If ws Is Nothing Then Set AsansorGetir = col: Exit Function

    Dim nSatir As Long
    nSatir = AsansorSatirBul(ws, nID)
    If nSatir = 0 Then Set AsansorGetir = col: Exit Function

    col.Add ws.Cells(nSatir, ASN_ID).Value,          "ID"
    col.Add ws.Cells(nSatir, ASN_BINA_ADI).Value,    "BinaAdi"
    col.Add ws.Cells(nSatir, ASN_ADRES).Value,        "Adres"
    col.Add ws.Cells(nSatir, ASN_ILCE).Value,         "Ilce"
    col.Add ws.Cells(nSatir, ASN_MALIK_ADI).Value,    "MalikAdi"
    col.Add ws.Cells(nSatir, ASN_MALIK_TEL).Value,    "MalikTel"
    col.Add ws.Cells(nSatir, ASN_MALIK_EMAIL).Value,  "MalikEmail"
    col.Add ws.Cells(nSatir, ASN_TESCIL_NO).Value,    "TescilNo"
    col.Add ws.Cells(nSatir, ASN_TIP).Value,           "Tip"
    col.Add ws.Cells(nSatir, ASN_KAPASITE).Value,      "Kapasite"
    col.Add ws.Cells(nSatir, ASN_DURAK_SAYISI).Value,  "DurakSayisi"
    col.Add ws.Cells(nSatir, ASN_MONTAJ_FIRMA).Value,  "MontajFirma"
    col.Add ws.Cells(nSatir, ASN_SERVIS_FIRMA).Value,  "ServisFirma"
    col.Add ws.Cells(nSatir, ASN_SON_PGD_TAR).Value,   "SonPGDTar"
    col.Add ws.Cells(nSatir, ASN_SON_PGD_SONUC).Value, "SonPGDSonuc"
    col.Add ws.Cells(nSatir, ASN_SON_PGD_BELGE).Value, "SonPGDBelge"
    col.Add ws.Cells(nSatir, ASN_SON_PGD_KURULUS).Value, "SonPGDKurulus"
    col.Add ws.Cells(nSatir, ASN_SON_PGD_TAR2).Value,  "SonrakiPGDTar"
    col.Add ws.Cells(nSatir, ASN_DURUM).Value,          "Durum"
    col.Add ws.Cells(nSatir, ASN_NOTLAR).Value,         "Notlar"

    Set AsansorGetir = col
End Function
