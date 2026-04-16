Attribute VB_Name = "modRaporlama"
'==============================================================================
' ASANSÖR PGD TAKİP SİSTEMİ - RAPORLAMA MODÜLü
' Modül : modRaporlama
' Açıklama: PGD durum özeti, ilçe bazlı istatistik, gecikmiş/yaklaşan
'           listeler ve Excel raporlarının üretilmesi.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' GİRİŞ SAYFASINI GÜNCELLE
' Dashboard'daki özet sayıları günceller. İş kitabı açıldığında da çalışır.
'------------------------------------------------------------------------------
Public Sub GirisSayfasiGuncelle()
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_GIRIS)
    If ws Is Nothing Then Exit Sub

    Dim wsA As Worksheet
    Set wsA = SayfaBul(SHT_ASANSORLER)
    If wsA Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    SayfaKorumaKapat ws

    Dim nToplamA As Long    ' Toplam asansör
    Dim nAktif As Long      ' Aktif
    Dim nMuhurlu As Long    ' Mühürlü
    Dim nDevreDisi As Long  ' Devre dışı
    Dim nGecikmiş As Long   ' PGD gecikmiş
    Dim nKritik As Long     ' 0-15 gün
    Dim nUyari As Long      ' 16-60 gün
    Dim nUygun As Long      ' >60 gün, PGD güncel

    Dim nSon As Long, i As Long
    nSon = SonSatir(wsA, 1)

    For i = 2 To nSon
        If wsA.Cells(i, ASN_ID).Value <> "" Then
            nToplamA = nToplamA + 1
            Dim sDurum As String
            sDurum = wsA.Cells(i, ASN_DURUM).Value
            Select Case sDurum
                Case DURUM_AKTIF:      nAktif = nAktif + 1
                Case DURUM_MUHURLU:    nMuhurlu = nMuhurlu + 1
                Case DURUM_DEVRE_DISI: nDevreDisi = nDevreDisi + 1
                Case Else:             nAktif = nAktif + 1
            End Select

            If sDurum = DURUM_AKTIF Then
                Dim dT As Variant
                dT = wsA.Cells(i, ASN_SON_PGD_TAR2).Value
                If IsEmpty(dT) Or Not IsDate(dT) Then
                    nGecikmiş = nGecikmiş + 1
                ElseIf CDate(dT) < Date Then
                    nGecikmiş = nGecikmiş + 1
                ElseIf CDate(dT) - Date <= PGD_KRITIK_ESIGI Then
                    nKritik = nKritik + 1
                ElseIf CDate(dT) - Date <= PGD_UYARI_ESIGI Then
                    nUyari = nUyari + 1
                Else
                    nUygun = nUygun + 1
                End If
            End If
        End If
    Next i

    ' GİRİŞ sayfasındaki belirli hücrelere yaz
    ' (Kurulum sırasında bu hücreler oluşturulur)
    ws.Range("C5").Value = nToplamA
    ws.Range("C6").Value = nAktif
    ws.Range("C7").Value = nMuhurlu
    ws.Range("C8").Value = nDevreDisi
    ws.Range("C10").Value = nGecikmiş
    ws.Range("C11").Value = nKritik
    ws.Range("C12").Value = nUyari
    ws.Range("C13").Value = nUygun
    ws.Range("C15").Value = TarihStr(Now) & " " & Format(Now, "HH:MM")

    ' Renklendirme
    ws.Range("C10").Interior.Color = IIf(nGecikmiş > 0, RENK_TEHLIKE, RENK_UYGUN)
    ws.Range("C10").Font.Color = IIf(nGecikmiş > 0, RENK_BASLIK_YAZ, RENK_BASLIK_YAZ)
    ws.Range("C11").Interior.Color = IIf(nKritik > 0, RENK_TEHLIKE, RENK_UYGUN)
    ws.Range("C11").Font.Color = RENK_BASLIK_YAZ
    ws.Range("C12").Interior.Color = IIf(nUyari > 0, RENK_UYARI, RENK_UYGUN)
    ws.Range("C12").Font.Color = RENK_BASLIK_YAZ

    SayfaKorumaAc ws
    Application.ScreenUpdating = True
End Sub

'------------------------------------------------------------------------------
' DURUM RAPORU ÜRETİMİ
' RAPORLAR sayfasına kapsamlı özet istatistik raporu yazar.
'------------------------------------------------------------------------------
Public Sub DurumRaporuUret()
    Dim wsR As Worksheet
    Set wsR = SayfaBul(SHT_RAPORLAR)
    If wsR Is Nothing Then Exit Sub

    Dim wsA As Worksheet
    Set wsA = SayfaBul(SHT_ASANSORLER)
    If wsA Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    SayfaKorumaKapat wsR
    wsR.Cells.Clear

    Dim nSatir As Long
    nSatir = 1

    ' Başlık
    With wsR.Cells(nSatir, 1)
        .Value = AyarOku("KURUM_ADI") & " - ASANSÖR PGD DURUM RAPORU"
        .Font.Bold = True
        .Font.Size = 14
    End With
    wsR.Range("A1:F1").Merge
    wsR.Range("A1:F1").HorizontalAlignment = xlCenter

    nSatir = 2
    wsR.Cells(nSatir, 1).Value = "Rapor Tarihi: " & TarihStr(Now) & " " & Format(Now, "HH:MM")
    wsR.Range("A2:F2").Merge
    nSatir = 4

    ' --- GENEL İSTATİSTİKLER ---
    wsR.Cells(nSatir, 1).Value = "GENEL İSTATİSTİKLER"
    BaslikStiliUygula wsR.Range("A" & nSatir & ":F" & nSatir)
    nSatir = nSatir + 1

    Dim ozet As Object
    Set ozet = HesaplaOzet(wsA)

    Dim satirData(10, 1) As Variant
    satirData(0, 0) = "Toplam Kayıtlı Asansör": satirData(0, 1) = ozet("toplam")
    satirData(1, 0) = "Aktif Asansör":          satirData(1, 1) = ozet("aktif")
    satirData(2, 0) = "Mühürlü Asansör":        satirData(2, 1) = ozet("muhurlu")
    satirData(3, 0) = "Devre Dışı Asansör":     satirData(3, 1) = ozet("devredisi")
    satirData(4, 0) = "--- PGD DURUM ---":       satirData(4, 1) = ""
    satirData(5, 0) = "PGD Gecikmiş":           satirData(5, 1) = ozet("gecikmiş")
    satirData(6, 0) = "PGD Kritik (0-15 gün)":  satirData(6, 1) = ozet("kritik")
    satirData(7, 0) = "PGD Yaklaşıyor (16-60 gün)": satirData(7, 1) = ozet("uyari")
    satirData(8, 0) = "PGD Güncel (>60 gün)":   satirData(8, 1) = ozet("uygun")
    satirData(9, 0) = "--- PGD KAYIT ---":       satirData(9, 1) = ""
    satirData(10, 0) = "Bu Yıl Yapılan PGD":     satirData(10, 1) = ozet("buYilPGD")

    Dim j As Long
    For j = 0 To 10
        wsR.Cells(nSatir, 1).Value = satirData(j, 0)
        wsR.Cells(nSatir, 2).Value = satirData(j, 1)
        If j = 4 Or j = 9 Then
            wsR.Cells(nSatir, 1).Font.Bold = True
        ElseIf j = 5 And CLng(satirData(j, 1)) > 0 Then
            wsR.Cells(nSatir, 2).Interior.Color = RENK_TEHLIKE
            wsR.Cells(nSatir, 2).Font.Color = RENK_BASLIK_YAZ
        End If
        nSatir = nSatir + 1
    Next j
    nSatir = nSatir + 1

    ' --- İLÇE BAZLI İSTATİSTİK ---
    wsR.Cells(nSatir, 1).Value = "İLÇE BAZLI DAĞILIM"
    BaslikStiliUygula wsR.Range("A" & nSatir & ":F" & nSatir)
    nSatir = nSatir + 1

    ' Başlık satırı
    wsR.Cells(nSatir, 1).Value = "İlçe"
    wsR.Cells(nSatir, 2).Value = "Toplam"
    wsR.Cells(nSatir, 3).Value = "Aktif"
    wsR.Cells(nSatir, 4).Value = "Mühürlü"
    wsR.Cells(nSatir, 5).Value = "Gecikmiş"
    wsR.Cells(nSatir, 6).Value = "PGD Güncel"
    BaslikStiliUygula wsR.Range("A" & nSatir & ":F" & nSatir)
    nSatir = nSatir + 1

    IlceBazliIstatistik wsA, wsR, nSatir
    nSatir = SonSatir(wsR, 1) + 2

    ' --- GECİKMİŞ ASANSÖRLER LİSTESİ ---
    wsR.Cells(nSatir, 1).Value = "GECİKMİŞ ASANSÖRLER (PGD SÜRESİ DOLMUŞ)"
    BaslikStiliUygula wsR.Range("A" & nSatir & ":F" & nSatir)
    nSatir = nSatir + 1

    wsR.Cells(nSatir, 1).Value = "ID"
    wsR.Cells(nSatir, 2).Value = "Bina Adı"
    wsR.Cells(nSatir, 3).Value = "İlçe"
    wsR.Cells(nSatir, 4).Value = "Tescil No"
    wsR.Cells(nSatir, 5).Value = "Son PGD"
    wsR.Cells(nSatir, 6).Value = "Gecikme"
    BaslikStiliUygula wsR.Range("A" & nSatir & ":F" & nSatir)
    nSatir = nSatir + 1

    GecikmisSatirlar wsA, wsR, nSatir

    ' Sütun genişliklerini ayarla
    wsR.Columns("A:F").AutoFit
    wsR.Columns("B").ColumnWidth = 30

    SayfaKorumaAc wsR
    Application.ScreenUpdating = True

    wsR.Activate
    MsgBox "Durum raporu oluşturuldu.", vbInformation, UYGULAMA_ADI
End Sub

'------------------------------------------------------------------------------
' ÖZET HESAPLA (iç kullanım)
'------------------------------------------------------------------------------
Private Function HesaplaOzet(ByVal wsA As Worksheet) As Object
    Dim col As New Collection

    Dim nSon As Long, i As Long
    nSon = SonSatir(wsA, 1)

    Dim nToplam As Long, nAktif As Long, nMuhurlu As Long, nDevreDisi As Long
    Dim nGecikmiş As Long, nKritik As Long, nUyari As Long, nUygun As Long
    Dim nBuYilPGD As Long

    For i = 2 To nSon
        If wsA.Cells(i, ASN_ID).Value <> "" Then
            nToplam = nToplam + 1
            Select Case wsA.Cells(i, ASN_DURUM).Value
                Case DURUM_AKTIF:      nAktif = nAktif + 1
                Case DURUM_MUHURLU:    nMuhurlu = nMuhurlu + 1
                Case DURUM_DEVRE_DISI: nDevreDisi = nDevreDisi + 1
                Case Else:             nAktif = nAktif + 1
            End Select

            If wsA.Cells(i, ASN_DURUM).Value = DURUM_AKTIF Then
                Dim dT As Variant
                dT = wsA.Cells(i, ASN_SON_PGD_TAR2).Value
                If IsEmpty(dT) Or Not IsDate(dT) Or CDate(dT) < Date Then
                    nGecikmiş = nGecikmiş + 1
                ElseIf CDate(dT) - Date <= PGD_KRITIK_ESIGI Then
                    nKritik = nKritik + 1
                ElseIf CDate(dT) - Date <= PGD_UYARI_ESIGI Then
                    nUyari = nUyari + 1
                Else
                    nUygun = nUygun + 1
                End If
            End If

            ' Bu yıl PGD yapılmış mı?
            Dim dPGD As Variant
            dPGD = wsA.Cells(i, ASN_SON_PGD_TAR).Value
            If IsDate(dPGD) Then
                If Year(CDate(dPGD)) = Year(Date) Then
                    nBuYilPGD = nBuYilPGD + 1
                End If
            End If
        End If
    Next i

    col.Add nToplam,    "toplam"
    col.Add nAktif,     "aktif"
    col.Add nMuhurlu,   "muhurlu"
    col.Add nDevreDisi, "devredisi"
    col.Add nGecikmiş,  "gecikmiş"
    col.Add nKritik,    "kritik"
    col.Add nUyari,     "uyari"
    col.Add nUygun,     "uygun"
    col.Add nBuYilPGD,  "buYilPGD"

    Set HesaplaOzet = col
End Function

'------------------------------------------------------------------------------
' İLÇE BAZLI İSTATİSTİK (iç kullanım)
'------------------------------------------------------------------------------
Private Sub IlceBazliIstatistik(ByVal wsA As Worksheet, ByVal wsR As Worksheet, ByVal nBasSatir As Long)
    ' Benzersiz ilçeleri bul
    Dim ilceler() As String
    ReDim ilceler(0)
    Dim nIlceCount As Long
    nIlceCount = 0

    Dim nSonA As Long, i As Long
    nSonA = SonSatir(wsA, 1)

    For i = 2 To nSonA
        If wsA.Cells(i, ASN_ID).Value <> "" Then
            Dim sIlce As String
            sIlce = Trim(wsA.Cells(i, ASN_ILCE).Value)
            If sIlce = "" Then sIlce = "(İlçe Girilmemiş)"
            Dim bBulundu As Boolean
            bBulundu = False
            Dim k As Long
            For k = 0 To nIlceCount - 1
                If ilceler(k) = sIlce Then bBulundu = True: Exit For
            Next k
            If Not bBulundu Then
                ReDim Preserve ilceler(nIlceCount)
                ilceler(nIlceCount) = sIlce
                nIlceCount = nIlceCount + 1
            End If
        End If
    Next i

    ' Her ilçe için istatistik hesapla
    Dim nSatir As Long
    nSatir = nBasSatir
    Dim m As Long
    For m = 0 To nIlceCount - 1
        Dim nToplam As Long, nAktif As Long, nMuhurlu As Long, nGecikmiş As Long, nUygun As Long
        nToplam = 0: nAktif = 0: nMuhurlu = 0: nGecikmiş = 0: nUygun = 0
        For i = 2 To nSonA
            If wsA.Cells(i, ASN_ID).Value <> "" Then
                Dim sIlce2 As String
                sIlce2 = Trim(wsA.Cells(i, ASN_ILCE).Value)
                If sIlce2 = "" Then sIlce2 = "(İlçe Girilmemiş)"
                If sIlce2 = ilceler(m) Then
                    nToplam = nToplam + 1
                    If wsA.Cells(i, ASN_DURUM).Value = DURUM_AKTIF Then nAktif = nAktif + 1
                    If wsA.Cells(i, ASN_DURUM).Value = DURUM_MUHURLU Then nMuhurlu = nMuhurlu + 1
                    Dim dT2 As Variant
                    dT2 = wsA.Cells(i, ASN_SON_PGD_TAR2).Value
                    If IsEmpty(dT2) Or Not IsDate(dT2) Or CDate(dT2) < Date Then
                        nGecikmiş = nGecikmiş + 1
                    Else
                        nUygun = nUygun + 1
                    End If
                End If
            End If
        Next i
        wsR.Cells(nSatir, 1).Value = ilceler(m)
        wsR.Cells(nSatir, 2).Value = nToplam
        wsR.Cells(nSatir, 3).Value = nAktif
        wsR.Cells(nSatir, 4).Value = nMuhurlu
        wsR.Cells(nSatir, 5).Value = nGecikmiş
        wsR.Cells(nSatir, 6).Value = nUygun
        If nGecikmiş > 0 Then
            wsR.Cells(nSatir, 5).Interior.Color = RENK_TEHLIKE
            wsR.Cells(nSatir, 5).Font.Color = RENK_BASLIK_YAZ
        End If
        wsR.Rows(nSatir).RowHeight = 18
        nSatir = nSatir + 1
    Next m
End Sub

'------------------------------------------------------------------------------
' GECİKMİŞ SATIRLAR (iç kullanım)
'------------------------------------------------------------------------------
Private Sub GecikmisSatirlar(ByVal wsA As Worksheet, ByVal wsR As Worksheet, ByVal nBasSatir As Long)
    Dim nSonA As Long, i As Long
    nSonA = SonSatir(wsA, 1)
    Dim nSatir As Long
    nSatir = nBasSatir
    For i = 2 To nSonA
        If wsA.Cells(i, ASN_ID).Value <> "" Then
            If wsA.Cells(i, ASN_DURUM).Value = DURUM_AKTIF Then
                Dim dT As Variant
                dT = wsA.Cells(i, ASN_SON_PGD_TAR2).Value
                If IsEmpty(dT) Or Not IsDate(dT) Or CDate(dT) < Date Then
                    Dim nGec As Long
                    If IsDate(dT) Then
                        nGec = Date - CDate(dT)
                    Else
                        nGec = 9999
                    End If
                    wsR.Cells(nSatir, 1).Value = wsA.Cells(i, ASN_ID).Value
                    wsR.Cells(nSatir, 2).Value = wsA.Cells(i, ASN_BINA_ADI).Value
                    wsR.Cells(nSatir, 3).Value = wsA.Cells(i, ASN_ILCE).Value
                    wsR.Cells(nSatir, 4).Value = wsA.Cells(i, ASN_TESCIL_NO).Value
                    wsR.Cells(nSatir, 5).Value = wsA.Cells(i, ASN_SON_PGD_TAR).Value
                    wsR.Cells(nSatir, 5).NumberFormat = "DD.MM.YYYY"
                    If nGec = 9999 Then
                        wsR.Cells(nSatir, 6).Value = "Tarih girilmemiş"
                    Else
                        wsR.Cells(nSatir, 6).Value = nGec & " gün"
                    End If
                    wsR.Rows(nSatir).Interior.Color = RGB(255, 220, 220)
                    wsR.Rows(nSatir).RowHeight = 18
                    nSatir = nSatir + 1
                End If
            End If
        End If
    Next i
    If nSatir = nBasSatir Then
        wsR.Cells(nSatir, 1).Value = "Gecikmiş asansör bulunmamaktadır."
        wsR.Cells(nSatir, 1).Font.Bold = True
        wsR.Cells(nSatir, 1).Interior.Color = RGB(204, 255, 204)
    End If
End Sub

'------------------------------------------------------------------------------
' TOPLU PGD BİLDİRİMİ
' Gecikmiş veya yaklaşan tüm asansörler için toplu yazı üretimi.
'------------------------------------------------------------------------------
Public Sub TopluPGDBildirimi()
    Dim wsA As Worksheet
    Set wsA = SayfaBul(SHT_ASANSORLER)
    If wsA Is Nothing Then Exit Sub

    Dim nGun As String
    nGun = Trim(InputBox("Kaç gün içinde PGD'si dolacak asansörler dahil edilsin?" & vbCrLf & _
                         "(0 = sadece gecikmiş, 60 = 60 gün içinde dolacaklar dahil)", _
                         UYGULAMA_ADI & " - Toplu Bildirim", "30"))
    If Not IsNumeric(nGun) Then Exit Sub
    Dim nEsik As Long
    nEsik = CLng(nGun)

    Dim nSayac As Long
    nSayac = 0
    Dim nSon As Long, i As Long
    nSon = SonSatir(wsA, 1)

    Dim sSoru As String
    sSoru = "Bildirimleri aşağıdaki kriterlere göre üretmek istiyorsunuz:" & vbCrLf & _
            "- Gecikmiş PGD + " & nEsik & " gün içinde dolacaklar" & vbCrLf & vbCrLf & _
            "Toplam kaç asansör için yazı üretileceği hesaplanıyor..."

    For i = 2 To nSon
        If wsA.Cells(i, ASN_ID).Value <> "" Then
            If wsA.Cells(i, ASN_DURUM).Value = DURUM_AKTIF Then
                Dim dT As Variant
                dT = wsA.Cells(i, ASN_SON_PGD_TAR2).Value
                If IsEmpty(dT) Or Not IsDate(dT) Or CDate(dT) - Date <= nEsik Then
                    nSayac = nSayac + 1
                End If
            End If
        End If
    Next i

    If nSayac = 0 Then
        MsgBox "Belirtilen kriterlere uyan asansör bulunamadı.", vbInformation, UYGULAMA_ADI
        Exit Sub
    End If

    If Not Onayla("Toplam " & nSayac & " asansör için PGD Bildirimi üretilecek." & vbCrLf & _
                  "Devam etmek istiyor musunuz? (İşlem birkaç dakika sürebilir)") Then Exit Sub

    Dim nUretilen As Long
    nUretilen = 0
    For i = 2 To nSon
        If wsA.Cells(i, ASN_ID).Value <> "" Then
            If wsA.Cells(i, ASN_DURUM).Value = DURUM_AKTIF Then
                Dim dT2 As Variant
                dT2 = wsA.Cells(i, ASN_SON_PGD_TAR2).Value
                If IsEmpty(dT2) Or Not IsDate(dT2) Or CDate(dT2) - Date <= nEsik Then
                    Dim nID As Long
                    nID = wsA.Cells(i, ASN_ID).Value
                    On Error Resume Next
                    PGDBildirimiUret nID
                    On Error GoTo 0
                    nUretilen = nUretilen + 1
                End If
            End If
        End If
    Next i

    MsgBox "Toplu bildirim tamamlandı." & vbCrLf & _
           "Üretilen yazı sayısı: " & nUretilen, vbInformation, UYGULAMA_ADI
End Sub
