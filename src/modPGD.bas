Attribute VB_Name = "modPGD"
'==============================================================================
' ASANSÖR PGD TAKİP SİSTEMİ - PGD KAYIT VE TAKİP YÖNETİMİ
' Modül : modPGD
' Açıklama: Periyodik Güvenlik Denetimi (PGD) kayıtlarının yönetimi;
'           yeni denetim kaydı girişi, takip sayfasının güncellenmesi
'           ve gecikmiş/yaklaşan denetimlerin listelenmesi.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' YENİ PGD KAYDI EKLE
' Mevcut bir asansöre yeni PGD denetim sonucunu kaydeder.
' Aynı zamanda ASANSÖRLER sayfasındaki son PGD bilgilerini günceller.
'------------------------------------------------------------------------------
Public Function PGDKayitEkle(Optional ByVal nAsansorID As Long = 0) As Long
    PGDKayitEkle = 0

    Dim wsA As Worksheet, wsP As Worksheet
    Set wsA = SayfaBul(SHT_ASANSORLER)
    Set wsP = SayfaBul(SHT_PGD_KAYIT)
    If wsA Is Nothing Or wsP Is Nothing Then
        MsgBox "Gerekli sayfalar bulunamadı!", vbCritical, UYGULAMA_ADI
        Exit Function
    End If

    ' Asansör ID seç
    If nAsansorID = 0 Then
        Dim sID As String
        sID = Trim(InputBox("PGD kaydı eklenecek Asansörün Kayıt ID'sini girin:", UYGULAMA_ADI))
        If sID = "" Or Not IsNumeric(sID) Then Exit Function
        nAsansorID = CLng(sID)
    End If

    ' Asansör bilgilerini getir
    Dim colA As Collection
    Set colA = AsansorGetir(nAsansorID)
    If colA.Count = 0 Then
        MsgBox "ID=" & nAsansorID & " olan asansör bulunamadı!", vbExclamation, UYGULAMA_ADI
        Exit Function
    End If

    ' Kullanıcıya asansör bilgisini göster
    MsgBox "Asansör Bilgisi:" & vbCrLf & vbCrLf & _
           "Bina       : " & colA("BinaAdi") & vbCrLf & _
           "Tescil No  : " & colA("TescilNo") & vbCrLf & _
           "Adres      : " & colA("Adres") & vbCrLf & _
           "Mevcut PGD : " & TarihStr(colA("SonPGDTar")) & vbCrLf & vbCrLf & _
           "Bu asansöre yeni PGD kaydı ekleyeceksiniz.", vbInformation, UYGULAMA_ADI

    ' PGD Bilgilerini al
    Dim sDenetimTar As String
    Dim sKurulus As String
    Dim sDenetimNo As String
    Dim sSonuc As String
    Dim sUygunsuzSay As String
    Dim sUygunsuzNot As String
    Dim sBelgeTar As String

    sDenetimTar = Trim(InputBox("Denetim Tarihi (GG.AA.YYYY):", UYGULAMA_ADI & " - PGD Kaydı", _
                                Format(Date, "DD.MM.YYYY")))
    If Not IsDate(sDenetimTar) Then
        MsgBox "Geçersiz denetim tarihi! Kayıt iptal edildi.", vbExclamation, UYGULAMA_ADI
        Exit Function
    End If

    sKurulus = Trim(InputBox("Denetim Kuruluşu (A-tipi muayene kuruluşu adı):", _
                             UYGULAMA_ADI & " - PGD Kaydı", colA("SonPGDKurulus")))
    If sKurulus = "" Then
        MsgBox "Denetim kuruluşu zorunludur!", vbExclamation, UYGULAMA_ADI
        Exit Function
    End If

    sDenetimNo = Trim(InputBox("Denetim / Belge No:", UYGULAMA_ADI & " - PGD Kaydı", ""))
    sSonuc = PGDSonucSecimYap()

    If sSonuc = SONUC_UYGUNSUZ Or sSonuc = SONUC_KISMI Then
        sUygunsuzSay = Trim(InputBox("Uygunsuzluk Sayısı:", UYGULAMA_ADI & " - PGD Kaydı", "0"))
        sUygunsuzNot = Trim(InputBox("Uygunsuzluk Notları (kısa açıklama):", _
                                     UYGULAMA_ADI & " - PGD Kaydı", ""))
    End If

    sBelgeTar = Trim(InputBox("Belge/Rapor Düzenleme Tarihi (GG.AA.YYYY, boş=denetim tarihi):", _
                              UYGULAMA_ADI & " - PGD Kaydı", sDenetimTar))
    If sBelgeTar = "" Or Not IsDate(sBelgeTar) Then sBelgeTar = sDenetimTar

    ' --- PGD_KAYITLAR sayfasına yaz ---
    Dim nPGDID As Long
    Dim nPGDSatir As Long
    nPGDID = YeniID(wsP)
    nPGDSatir = SonSatir(wsP, 1) + 1

    SayfaKorumaKapat wsP
    With wsP
        .Cells(nPGDSatir, PGD_ID).Value = nPGDID
        .Cells(nPGDSatir, PGD_ASANSOR_ID).Value = nAsansorID
        .Cells(nPGDSatir, PGD_BINA_ADI).Value = colA("BinaAdi")
        .Cells(nPGDSatir, PGD_TESCIL_NO).Value = colA("TescilNo")
        .Cells(nPGDSatir, PGD_DENETIM_TAR).Value = CDate(sDenetimTar)
        .Cells(nPGDSatir, PGD_DENETIM_TAR).NumberFormat = "DD.MM.YYYY"
        .Cells(nPGDSatir, PGD_KURULUS).Value = sKurulus
        .Cells(nPGDSatir, PGD_DENETIM_NO).Value = sDenetimNo
        .Cells(nPGDSatir, PGD_SONUC).Value = sSonuc
        If IsNumeric(sUygunsuzSay) Then
            .Cells(nPGDSatir, PGD_UYGUNSUZ_SAY).Value = CInt(sUygunsuzSay)
        End If
        .Cells(nPGDSatir, PGD_UYGUNSUZ_NOT).Value = sUygunsuzNot
        .Cells(nPGDSatir, PGD_BELGE_TAR).Value = CDate(sBelgeTar)
        .Cells(nPGDSatir, PGD_BELGE_TAR).NumberFormat = "DD.MM.YYYY"
        .Cells(nPGDSatir, PGD_KAYIT_TAR).Value = Now
        .Cells(nPGDSatir, PGD_KAYIT_TAR).NumberFormat = "DD.MM.YYYY HH:MM"
    End With
    ' Satır rengi - sonuca göre
    PGDSatiriRenklendir wsP, nPGDSatir, sSonuc
    SayfaKorumaAc wsP

    ' --- ASANSÖRLER sayfasındaki son PGD bilgilerini güncelle ---
    Dim nASatir As Long
    nASatir = AsansorSatiriBulPublic(wsA, nAsansorID)
    If nASatir > 0 Then
        SayfaKorumaKapat wsA
        With wsA
            .Cells(nASatir, ASN_SON_PGD_TAR).Value = CDate(sDenetimTar)
            .Cells(nASatir, ASN_SON_PGD_TAR).NumberFormat = "DD.MM.YYYY"
            .Cells(nASatir, ASN_SON_PGD_SONUC).Value = sSonuc
            .Cells(nASatir, ASN_SON_PGD_BELGE).Value = sDenetimNo
            .Cells(nASatir, ASN_SON_PGD_KURULUS).Value = sKurulus
            .Cells(nASatir, ASN_SON_PGD_TAR2).Value = CDate(sDenetimTar) + PGD_SURE_GUN
            .Cells(nASatir, ASN_SON_PGD_TAR2).NumberFormat = "DD.MM.YYYY"
            ' Mühürlü ise durumu güncelle
            If sSonuc = SONUC_UYGUNSUZ Then
                If Onayla("Sonuç 'Uygunsuz' seçildi. Asansörü MÜHÜRLÜ olarak işaretlemek ister misiniz?") Then
                    .Cells(nASatir, ASN_DURUM).Value = DURUM_MUHURLU
                End If
            ElseIf sSonuc = SONUC_UYGUN Then
                If .Cells(nASatir, ASN_DURUM).Value = DURUM_MUHURLU Then
                    If Onayla("Asansör daha önce mühürlüydü. Durumu 'Aktif' olarak güncellensin mi?") Then
                        .Cells(nASatir, ASN_DURUM).Value = DURUM_AKTIF
                    End If
                End If
            End If
        End With
        AsansorleriSatirGuncelle wsA, nASatir
        SayfaKorumaAc wsA
    End If

    MsgBox "PGD kaydı başarıyla eklendi!" & vbCrLf & _
           "PGD Kayıt ID : " & nPGDID & vbCrLf & _
           "Sonuç        : " & sSonuc & vbCrLf & _
           "Sonraki PGD  : " & TarihStr(CDate(sDenetimTar) + PGD_SURE_GUN), _
           vbInformation, UYGULAMA_ADI

    PGDKayitEkle = nPGDID
End Function

'------------------------------------------------------------------------------
' TAKVİM SAYFASINI GÜNCELLE
' Tüm aktif asansörlerin PGD durumunu TAKVİM sayfasına yazar.
'------------------------------------------------------------------------------
Public Sub TakvimGuncelle()
    Dim wsA As Worksheet, wsT As Worksheet
    Set wsA = SayfaBul(SHT_ASANSORLER)
    Set wsT = SayfaBul(SHT_TAKVIM)
    If wsA Is Nothing Or wsT Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    SayfaKorumaKapat wsT

    ' Takvim sayfasını temizle (başlık satırı hariç)
    Dim nSonT As Long
    nSonT = SonSatir(wsT, 1)
    If nSonT > 1 Then
        wsT.Rows("2:" & nSonT).ClearContents
        wsT.Rows("2:" & nSonT).Interior.ColorIndex = xlNone
    End If

    ' Asansörlerden veri al ve sırala (en yakın PGD tarihli önce)
    Dim nSonA As Long
    nSonA = SonSatir(wsA, 1)
    If nSonA < 2 Then
        SayfaKorumaAc wsT
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Geçici dizi ile sıralama yap
    Dim nTakvimSatir As Long
    nTakvimSatir = 2

    ' Önce gecikmiş kayıtları yaz
    Dim i As Long
    For i = 2 To nSonA
        If wsA.Cells(i, ASN_ID).Value <> "" Then
            If wsA.Cells(i, ASN_DURUM).Value <> DURUM_DEVRE_DISI Then
                Dim dSonraki As Variant
                dSonraki = wsA.Cells(i, ASN_SON_PGD_TAR2).Value
                Dim nKalan As Long
                If IsDate(dSonraki) Then
                    nKalan = CDate(dSonraki) - Date
                Else
                    nKalan = -9999 ' Tarih girilmemiş = gecikmiş say
                End If
                If nKalan < 0 Then
                    TakvimSatirYaz wsT, wsA, i, nTakvimSatir, nKalan
                    nTakvimSatir = nTakvimSatir + 1
                End If
            End If
        End If
    Next i

    ' Sonra yaklaşan (0-60 gün) kayıtları yaz
    For i = 2 To nSonA
        If wsA.Cells(i, ASN_ID).Value <> "" Then
            If wsA.Cells(i, ASN_DURUM).Value <> DURUM_DEVRE_DISI Then
                Dim dSonraki2 As Variant
                dSonraki2 = wsA.Cells(i, ASN_SON_PGD_TAR2).Value
                If IsDate(dSonraki2) Then
                    Dim nKalan2 As Long
                    nKalan2 = CDate(dSonraki2) - Date
                    If nKalan2 >= 0 And nKalan2 <= PGD_UYARI_ESIGI Then
                        TakvimSatirYaz wsT, wsA, i, nTakvimSatir, nKalan2
                        nTakvimSatir = nTakvimSatir + 1
                    End If
                End If
            End If
        End If
    Next i

    ' Son olarak uygun durumda olanları yaz (>60 gün)
    For i = 2 To nSonA
        If wsA.Cells(i, ASN_ID).Value <> "" Then
            If wsA.Cells(i, ASN_DURUM).Value <> DURUM_DEVRE_DISI Then
                Dim dSonraki3 As Variant
                dSonraki3 = wsA.Cells(i, ASN_SON_PGD_TAR2).Value
                If IsDate(dSonraki3) Then
                    If CDate(dSonraki3) - Date > PGD_UYARI_ESIGI Then
                        TakvimSatirYaz wsT, wsA, i, nTakvimSatir, CLng(CDate(dSonraki3) - Date)
                        nTakvimSatir = nTakvimSatir + 1
                    End If
                End If
            End If
        End If
    Next i

    SayfaKorumaAc wsT
    Application.ScreenUpdating = True

    MsgBox "Takvim sayfası güncellendi. Toplam " & (nTakvimSatir - 2) & " asansör listelendi.", _
           vbInformation, UYGULAMA_ADI
End Sub

'------------------------------------------------------------------------------
' GECİKMİŞ ASANSÖR LİSTESİ
' PGD tarihi geçmiş asansörlerin sayısını ve listesini döndürür.
'------------------------------------------------------------------------------
Public Function GecikmisMiktar() As Long
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_ASANSORLER)
    If ws Is Nothing Then GecikmisMiktar = 0: Exit Function

    Dim nSon As Long, i As Long, nSayac As Long
    nSon = SonSatir(ws, 1)
    nSayac = 0
    For i = 2 To nSon
        If ws.Cells(i, ASN_ID).Value <> "" Then
            If ws.Cells(i, ASN_DURUM).Value = DURUM_AKTIF Then
                Dim dT As Variant
                dT = ws.Cells(i, ASN_SON_PGD_TAR2).Value
                If IsEmpty(dT) Or Not IsDate(dT) Or CDate(dT) < Date Then
                    nSayac = nSayac + 1
                End If
            End If
        End If
    Next i
    GecikmisMiktar = nSayac
End Function

'------------------------------------------------------------------------------
' YAKLAŞAN PGD SAYISI
' Belirtilen gün içinde PGD'si dolacak aktif asansör sayısı.
'------------------------------------------------------------------------------
Public Function YaklasanMiktar(Optional ByVal nGun As Integer = 60) As Long
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_ASANSORLER)
    If ws Is Nothing Then YaklasanMiktar = 0: Exit Function

    Dim nSon As Long, i As Long, nSayac As Long
    nSon = SonSatir(ws, 1)
    nSayac = 0
    For i = 2 To nSon
        If ws.Cells(i, ASN_ID).Value <> "" Then
            If ws.Cells(i, ASN_DURUM).Value = DURUM_AKTIF Then
                Dim dT As Variant
                dT = ws.Cells(i, ASN_SON_PGD_TAR2).Value
                If IsDate(dT) Then
                    Dim nKalan As Long
                    nKalan = CDate(dT) - Date
                    If nKalan >= 0 And nKalan <= nGun Then
                        nSayac = nSayac + 1
                    End If
                End If
            End If
        End If
    Next i
    YaklasanMiktar = nSayac
End Function

'------------------------------------------------------------------------------
' YARDIMCI: Takvim satırına veri yazar
'------------------------------------------------------------------------------
Private Sub TakvimSatirYaz(ByVal wsT As Worksheet, ByVal wsA As Worksheet, _
                            ByVal nASatir As Long, ByVal nTSatir As Long, ByVal nKalanGun As Long)
    With wsT
        .Cells(nTSatir, 1).Value = wsA.Cells(nASatir, ASN_ID).Value
        .Cells(nTSatir, 2).Value = wsA.Cells(nASatir, ASN_BINA_ADI).Value
        .Cells(nTSatir, 3).Value = wsA.Cells(nASatir, ASN_ADRES).Value
        .Cells(nTSatir, 4).Value = wsA.Cells(nASatir, ASN_ILCE).Value
        .Cells(nTSatir, 5).Value = wsA.Cells(nASatir, ASN_TESCIL_NO).Value
        .Cells(nTSatir, 6).Value = wsA.Cells(nASatir, ASN_MALIK_ADI).Value
        .Cells(nTSatir, 7).Value = wsA.Cells(nASatir, ASN_MALIK_TEL).Value
        .Cells(nTSatir, 8).Value = wsA.Cells(nASatir, ASN_SON_PGD_TAR).Value
        .Cells(nTSatir, 8).NumberFormat = "DD.MM.YYYY"
        .Cells(nTSatir, 9).Value = wsA.Cells(nASatir, ASN_SON_PGD_TAR2).Value
        .Cells(nTSatir, 9).NumberFormat = "DD.MM.YYYY"
        .Cells(nTSatir, 10).Value = PGDDurumMetni(wsA.Cells(nASatir, ASN_SON_PGD_TAR2).Value)
        .Cells(nTSatir, 11).Value = wsA.Cells(nASatir, ASN_DURUM).Value

        ' Satırı renklendir
        Dim renkBG As Long
        renkBG = PGDDurumRengi(wsA.Cells(nASatir, ASN_SON_PGD_TAR2).Value)
        .Rows(nTSatir).Interior.Color = renkBG
        If renkBG <> RENK_ZEBRA1 And renkBG <> RENK_ZEBRA2 Then
            .Rows(nTSatir).Font.Color = RENK_BASLIK_YAZ
        Else
            .Rows(nTSatir).Font.Color = RGB(50, 50, 50)
        End If
        .Rows(nTSatir).RowHeight = 20
    End With
End Sub

'------------------------------------------------------------------------------
' YARDIMCI: ASANSÖRLER sayfasında ID'ye göre satır bul (Public erişim)
'------------------------------------------------------------------------------
Public Function AsansorSatiriBulPublic(ByVal ws As Worksheet, ByVal nID As Long) As Long
    Dim nSon As Long, i As Long
    nSon = SonSatir(ws, 1)
    For i = 2 To nSon
        If ws.Cells(i, ASN_ID).Value = nID Then
            AsansorSatiriBulPublic = i
            Exit Function
        End If
    Next i
    AsansorSatiriBulPublic = 0
End Function

'------------------------------------------------------------------------------
' YARDIMCI: ASANSÖRLER sayfasının tek satırını renklendir
'------------------------------------------------------------------------------
Public Sub AsansorleriSatirGuncelle(ByVal ws As Worksheet, ByVal nSatir As Long)
    Dim dSonraki As Variant
    dSonraki = ws.Cells(nSatir, ASN_SON_PGD_TAR2).Value
    ws.Cells(nSatir, ASN_SON_PGD_TAR2).Interior.Color = PGDDurumRengi(dSonraki)
    If PGDDurumRengi(dSonraki) <> RENK_ZEBRA1 And PGDDurumRengi(dSonraki) <> RENK_ZEBRA2 Then
        ws.Cells(nSatir, ASN_SON_PGD_TAR2).Font.Color = RENK_BASLIK_YAZ
    End If
End Sub

'------------------------------------------------------------------------------
' YARDIMCI: PGD Sonuç seçim menüsü
'------------------------------------------------------------------------------
Private Function PGDSonucSecimYap() As String
    Dim sSecenekler As String
    sSecenekler = "1 - " & SONUC_UYGUN & vbCrLf & _
                  "2 - " & SONUC_UYGUNSUZ & vbCrLf & _
                  "3 - " & SONUC_KISMI
    Dim sSec As String
    sSec = Trim(InputBox("PGD Denetim Sonucunu seçin:" & vbCrLf & vbCrLf & sSecenekler, _
                         UYGULAMA_ADI & " - PGD Sonucu", "1"))
    Select Case sSec
        Case "1": PGDSonucSecimYap = SONUC_UYGUN
        Case "2": PGDSonucSecimYap = SONUC_UYGUNSUZ
        Case "3": PGDSonucSecimYap = SONUC_KISMI
        Case Else: PGDSonucSecimYap = SONUC_UYGUN
    End Select
End Function

'------------------------------------------------------------------------------
' YARDIMCI: PGD_KAYITLAR satırını sonuca göre renklendir
'------------------------------------------------------------------------------
Private Sub PGDSatiriRenklendir(ByVal ws As Worksheet, ByVal nSatir As Long, ByVal sSonuc As String)
    Select Case sSonuc
        Case SONUC_UYGUN
            ws.Rows(nSatir).Interior.Color = RGB(204, 255, 204)  ' Açık yeşil
        Case SONUC_UYGUNSUZ
            ws.Rows(nSatir).Interior.Color = RGB(255, 204, 204)  ' Açık kırmızı
        Case SONUC_KISMI
            ws.Rows(nSatir).Interior.Color = RGB(255, 255, 204)  ' Açık sarı
        Case Else
            ws.Rows(nSatir).Interior.ColorIndex = xlNone
    End Select
    ws.Rows(nSatir).RowHeight = 20
End Sub
