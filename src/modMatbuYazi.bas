Attribute VB_Name = "modMatbuYazi"
'==============================================================================
' ASANSÖR PGD TAKİP SİSTEMİ - MATBU YAZI ÜRETME MODÜLü
' Modül : modMatbuYazi
' Açıklama: Resmi yazıların (PGD bildirimi, ihtar, mühürleme kararı,
'           uygunluk belgesi) Word belgesi olarak otomatik üretilmesi.
'           Microsoft Word ile COM otomasyonu kullanır (geç bağlama).
'
' Şablon değişken formatı: <<DEĞİŞKEN_ADI>>
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' ANA YAZI ÜRETME MENÜSÜ
' Kullanıcıya üretilecek yazı türünü seçtirir.
'------------------------------------------------------------------------------
Public Sub MatbuYaziUret()
    Dim sMenu As String
    sMenu = "Üretilecek yazı türünü seçin:" & vbCrLf & vbCrLf & _
            "1 - PGD Bildirimi" & vbCrLf & _
            "2 - İhtar / Uyarı Yazısı" & vbCrLf & _
            "3 - Mühürleme / Kapatma Kararı" & vbCrLf & _
            "4 - Uygunluk / İzin Belgesi" & vbCrLf & _
            "5 - Güvensiz Asansör İdari Yaptırım Üst Yazısı"

    Dim sSec As String
    sSec = Trim(InputBox(sMenu, UYGULAMA_ADI & " - Matbu Yazı Üret", "1"))

    Dim nAsansorID As Long
    nAsansorID = AsansorIDSec()
    If nAsansorID = 0 Then Exit Sub

    Select Case sSec
        Case "1": PGDBildirimiUret nAsansorID
        Case "2": IhtarYazisiUret nAsansorID
        Case "3": MuhürlemeKarariUret nAsansorID
        Case "4": UygunlukBelgesiUret nAsansorID
        Case "5": GuvensizAsansorYaptirimUret nAsansorID
        Case Else
            MsgBox "Geçersiz seçim.", vbExclamation, UYGULAMA_ADI
    End Select
End Sub

'------------------------------------------------------------------------------
' PGD BİLDİRİMİ ÜRETİMİ
' Malik/yöneticiye PGD yaptırma zorunluluğunu bildiren resmi yazı.
'------------------------------------------------------------------------------
Public Sub PGDBildirimiUret(ByVal nAsansorID As Long)
    Dim col As Collection
    Set col = AsansorGetir(nAsansorID)
    If col.Count = 0 Then
        MsgBox "Asansör bulunamadı!", vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If

    Dim sSablon As String
    sSablon = SablonOku("PGD_BILDIRIMI")
    If sSablon = "" Then
        MsgBox "PGD Bildirimi şablonu ŞABLONLAR sayfasında bulunamadı!", vbCritical, UYGULAMA_ADI
        Exit Sub
    End If

    ' Son PGD tarihini bul; yoksa bugünü kullan
    Dim sSonPGD As String
    sSonPGD = TarihStr(col("SonPGDTar"))
    If sSonPGD = "" Then sSonPGD = "---"

    ' Son tarihi kullanıcıdan al (önerilen: son PGD + 30 gün)
    Dim dOnerilen As Date
    If IsDate(col("SonPGDTar")) Then
        dOnerilen = CDate(col("SonPGDTar")) + 30
    Else
        dOnerilen = Date + 30
    End If
    Dim sSure As String
    sSure = Trim(InputBox("Malik'e tanınan süre (GG.AA.YYYY son tarih):", _
                          UYGULAMA_ADI & " - PGD Bildirimi", TarihStr(dOnerilen)))
    If Not IsDate(sSure) Then sSure = TarihStr(dOnerilen)

    ' Şablon değişkenlerini doldur
    Dim degMap() As String
    ReDim degMap(0 To 19, 0 To 1)
    KurumBilgileriniDoldur degMap
    degMap(8, 0) = "<<BINA_ADI>>":      degMap(8, 1) = col("BinaAdi")
    degMap(9, 0) = "<<ADRES>>":         degMap(9, 1) = col("Adres") & " " & col("Ilce")
    degMap(10, 0) = "<<MALIK_ADI>>":    degMap(10, 1) = col("MalikAdi")
    degMap(11, 0) = "<<TESCIL_NO>>":    degMap(11, 1) = col("TescilNo")
    degMap(12, 0) = "<<SON_PGD_TAR>>":  degMap(12, 1) = sSonPGD
    degMap(13, 0) = "<<SON_TAR>>":      degMap(13, 1) = sSure
    degMap(14, 0) = "<<ASANSOR_TIP>>":  degMap(14, 1) = col("Tip")
    degMap(15, 0) = "<<KAPASITE>>":     degMap(15, 1) = CStr(col("Kapasite")) & " kg"
    degMap(16, 0) = "<<SERVIS_FIRMA>>": degMap(16, 1) = col("ServisFirma")
    degMap(17, 0) = "<<DURUM>>":        degMap(17, 1) = col("Durum")
    degMap(18, 0) = "<<TARIH>>":        degMap(18, 1) = TarihStr(Date)
    degMap(19, 0) = "<<YIL>>":          degMap(19, 1) = Year(Date)

    Dim sIcerik As String
    sIcerik = DegiskenleriDoldur(sSablon, degMap)

    Dim sDosyaYolu As String
    sDosyaYolu = WordBelgesiOlustur(sIcerik, _
                 "PGD_Bildirimi_" & col("TescilNo") & "_" & Format(Date, "YYYYMMDD"))

    If sDosyaYolu <> "" Then
        MsgBox "PGD Bildirimi oluşturuldu:" & vbCrLf & sDosyaYolu, vbInformation, UYGULAMA_ADI
    End If
End Sub

'------------------------------------------------------------------------------
' İHTAR / UYARI YAZISI ÜRETİMİ
' Süresinde PGD yaptırmayanlara gönderilen ihtarlı yazı.
'------------------------------------------------------------------------------
Public Sub IhtarYazisiUret(ByVal nAsansorID As Long)
    Dim col As Collection
    Set col = AsansorGetir(nAsansorID)
    If col.Count = 0 Then
        MsgBox "Asansör bulunamadı!", vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If

    Dim sSablon As String
    sSablon = SablonOku("IHTAR_YAZISI")
    If sSablon = "" Then
        MsgBox "İhtar Yazısı şablonu ŞABLONLAR sayfasında bulunamadı!", vbCritical, UYGULAMA_ADI
        Exit Sub
    End If

    ' Gecikme hesabı
    Dim nGecikme As Long
    Dim sPGDSonTar As String
    nGecikme = 0
    sPGDSonTar = TarihStr(col("SonrakiPGDTar"))
    If IsDate(col("SonrakiPGDTar")) Then
        nGecikme = Date - CDate(col("SonrakiPGDTar"))
        If nGecikme < 0 Then nGecikme = 0
    End If

    ' İdari para cezası (kullanıcıdan al veya varsayılan)
    Dim sCeza As String
    sCeza = Trim(InputBox("Uygulanacak idari para cezası tutarı (TL):", _
                          UYGULAMA_ADI & " - İhtar Yazısı", ""))

    ' Süre
    Dim sSure As String
    sSure = Trim(InputBox("Uyum için verilen süre (gün):", _
                          UYGULAMA_ADI & " - İhtar Yazısı", "15"))
    If Not IsNumeric(sSure) Or CInt(sSure) < 1 Then sSure = "15"
    Dim dSure As Date
    dSure = Date + CInt(sSure)

    Dim degMap() As String
    ReDim degMap(0 To 21, 0 To 1)
    KurumBilgileriniDoldur degMap
    degMap(8, 0) = "<<BINA_ADI>>":       degMap(8, 1) = col("BinaAdi")
    degMap(9, 0) = "<<ADRES>>":          degMap(9, 1) = col("Adres") & " " & col("Ilce")
    degMap(10, 0) = "<<MALIK_ADI>>":     degMap(10, 1) = col("MalikAdi")
    degMap(11, 0) = "<<TESCIL_NO>>":     degMap(11, 1) = col("TescilNo")
    degMap(12, 0) = "<<SON_PGD_TAR>>":   degMap(12, 1) = TarihStr(col("SonPGDTar"))
    degMap(13, 0) = "<<GECIKME_GUN>>":   degMap(13, 1) = CStr(nGecikme) & " gün"
    degMap(14, 0) = "<<PGD_SON_TAR>>":   degMap(14, 1) = sPGDSonTar
    degMap(15, 0) = "<<CEZA_TUTARI>>":   degMap(15, 1) = sCeza & " TL"
    degMap(16, 0) = "<<SURE_GUN>>":      degMap(16, 1) = sSure & " gün"
    degMap(17, 0) = "<<SURE_SON_TAR>>":  degMap(17, 1) = TarihStr(dSure)
    degMap(18, 0) = "<<ASANSOR_TIP>>":   degMap(18, 1) = col("Tip")
    degMap(19, 0) = "<<SERVIS_FIRMA>>":  degMap(19, 1) = col("ServisFirma")
    degMap(20, 0) = "<<TARIH>>":         degMap(20, 1) = TarihStr(Date)
    degMap(21, 0) = "<<YIL>>":           degMap(21, 1) = Year(Date)

    Dim sIcerik As String
    sIcerik = DegiskenleriDoldur(sSablon, degMap)

    Dim sDosyaYolu As String
    sDosyaYolu = WordBelgesiOlustur(sIcerik, _
                 "Ihtar_" & col("TescilNo") & "_" & Format(Date, "YYYYMMDD"))

    If sDosyaYolu <> "" Then
        MsgBox "İhtar Yazısı oluşturuldu:" & vbCrLf & sDosyaYolu, vbInformation, UYGULAMA_ADI
    End If
End Sub

'------------------------------------------------------------------------------
' MÜHÜRLEME / KAPATMA KARARI ÜRETİMİ
'------------------------------------------------------------------------------
Public Sub MuhürlemeKarariUret(ByVal nAsansorID As Long)
    Dim col As Collection
    Set col = AsansorGetir(nAsansorID)
    If col.Count = 0 Then
        MsgBox "Asansör bulunamadı!", vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If

    Dim sSablon As String
    sSablon = SablonOku("MUHÜRLEME_KARARI")
    If sSablon = "" Then
        MsgBox "Mühürleme Kararı şablonu ŞABLONLAR sayfasında bulunamadı!", vbCritical, UYGULAMA_ADI
        Exit Sub
    End If

    Dim sGerekce As String
    sGerekce = Trim(InputBox("Mühürleme gerekçesi:" & vbCrLf & _
                             "(Örn: PGD süresi dolmuş ve yaptırılmamış / Denetimde uygunsuz bulunmuş)", _
                             UYGULAMA_ADI & " - Mühürleme Kararı", _
                             "Asansörün Periyodik Güvenlik Denetimi (PGD) süresi dolmuş olup " & _
                             "malik tarafından yaptırılmamıştır."))

    Dim sKararNo As String
    sKararNo = Trim(InputBox("Karar Numarası:", UYGULAMA_ADI & " - Mühürleme Kararı", _
                             Year(Date) & "/" & Format(Date, "MMDD")))

    ' Asansörü mühürlü olarak işaretle
    If Onayla("Asansör durumunu 'Mühürlü' olarak güncellensin mi?") Then
        Dim ws As Worksheet
        Set ws = SayfaBul(SHT_ASANSORLER)
        If Not ws Is Nothing Then
            Dim nSatir As Long
            nSatir = AsansorSatiriBulPublic(ws, nAsansorID)
            If nSatir > 0 Then
                SayfaKorumaKapat ws
                ws.Cells(nSatir, ASN_DURUM).Value = DURUM_MUHURLU
                SayfaKorumaAc ws
            End If
        End If
    End If

    Dim degMap() As String
    ReDim degMap(0 To 20, 0 To 1)
    KurumBilgileriniDoldur degMap
    degMap(8, 0) = "<<BINA_ADI>>":      degMap(8, 1) = col("BinaAdi")
    degMap(9, 0) = "<<ADRES>>":         degMap(9, 1) = col("Adres") & " " & col("Ilce")
    degMap(10, 0) = "<<MALIK_ADI>>":    degMap(10, 1) = col("MalikAdi")
    degMap(11, 0) = "<<TESCIL_NO>>":    degMap(11, 1) = col("TescilNo")
    degMap(12, 0) = "<<SON_PGD_TAR>>":  degMap(12, 1) = TarihStr(col("SonPGDTar"))
    degMap(13, 0) = "<<KARAR_NO>>":     degMap(13, 1) = sKararNo
    degMap(14, 0) = "<<GEREKCE>>":      degMap(14, 1) = sGerekce
    degMap(15, 0) = "<<ASANSOR_TIP>>":  degMap(15, 1) = col("Tip")
    degMap(16, 0) = "<<KAPASITE>>":     degMap(16, 1) = CStr(col("Kapasite")) & " kg"
    degMap(17, 0) = "<<SERVIS_FIRMA>>": degMap(17, 1) = col("ServisFirma")
    degMap(18, 0) = "<<TARIH>>":        degMap(18, 1) = TarihStr(Date)
    degMap(19, 0) = "<<YIL>>":          degMap(19, 1) = Year(Date)
    degMap(20, 0) = "<<DURAK_SAYISI>>": degMap(20, 1) = CStr(col("DurakSayisi"))

    Dim sIcerik As String
    sIcerik = DegiskenleriDoldur(sSablon, degMap)

    Dim sDosyaYolu As String
    sDosyaYolu = WordBelgesiOlustur(sIcerik, _
                 "Muhürleme_" & col("TescilNo") & "_" & Format(Date, "YYYYMMDD"))

    If sDosyaYolu <> "" Then
        MsgBox "Mühürleme Kararı oluşturuldu:" & vbCrLf & sDosyaYolu, vbInformation, UYGULAMA_ADI
    End If
End Sub

'------------------------------------------------------------------------------
' UYGUNLUK / İZİN BELGESİ ÜRETİMİ
' PGD'yi başarıyla tamamlayan asansörler için belge üretir.
'------------------------------------------------------------------------------
Public Sub UygunlukBelgesiUret(ByVal nAsansorID As Long)
    Dim col As Collection
    Set col = AsansorGetir(nAsansorID)
    If col.Count = 0 Then
        MsgBox "Asansör bulunamadı!", vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If

    If col("SonPGDSonuc") <> SONUC_UYGUN Then
        If Not Onayla("Son PGD sonucu '" & col("SonPGDSonuc") & "' olarak kayıtlı." & vbCrLf & _
                      "Yine de uygunluk belgesi üretmek istiyor musunuz?") Then Exit Sub
    End If

    Dim sSablon As String
    sSablon = SablonOku("UYGUNLUK_BELGESI")
    If sSablon = "" Then
        MsgBox "Uygunluk Belgesi şablonu ŞABLONLAR sayfasında bulunamadı!", vbCritical, UYGULAMA_ADI
        Exit Sub
    End If

    Dim sBelgeNo As String
    sBelgeNo = Trim(InputBox("Belge Numarası:", UYGULAMA_ADI & " - Uygunluk Belgesi", _
                             Year(Date) & "/" & Format(Date, "MMDD") & "/" & nAsansorID))

    Dim degMap() As String
    ReDim degMap(0 To 21, 0 To 1)
    KurumBilgileriniDoldur degMap
    degMap(8, 0) = "<<BINA_ADI>>":        degMap(8, 1) = col("BinaAdi")
    degMap(9, 0) = "<<ADRES>>":           degMap(9, 1) = col("Adres") & " " & col("Ilce")
    degMap(10, 0) = "<<MALIK_ADI>>":      degMap(10, 1) = col("MalikAdi")
    degMap(11, 0) = "<<TESCIL_NO>>":      degMap(11, 1) = col("TescilNo")
    degMap(12, 0) = "<<PGD_TAR>>":        degMap(12, 1) = TarihStr(col("SonPGDTar"))
    degMap(13, 0) = "<<PGD_BELGE_NO>>":   degMap(13, 1) = col("SonPGDBelge")
    degMap(14, 0) = "<<PGD_KURULUS>>":    degMap(14, 1) = col("SonPGDKurulus")
    degMap(15, 0) = "<<SONRAKI_PGD>>":    degMap(15, 1) = TarihStr(col("SonrakiPGDTar"))
    degMap(16, 0) = "<<ASANSOR_TIP>>":    degMap(16, 1) = col("Tip")
    degMap(17, 0) = "<<KAPASITE>>":       degMap(17, 1) = CStr(col("Kapasite")) & " kg"
    degMap(18, 0) = "<<DURAK_SAYISI>>":   degMap(18, 1) = CStr(col("DurakSayisi"))
    degMap(19, 0) = "<<BELGE_NO>>":       degMap(19, 1) = sBelgeNo
    degMap(20, 0) = "<<TARIH>>":          degMap(20, 1) = TarihStr(Date)
    degMap(21, 0) = "<<YIL>>":            degMap(21, 1) = Year(Date)

    Dim sIcerik As String
    sIcerik = DegiskenleriDoldur(sSablon, degMap)

    Dim sDosyaYolu As String
    sDosyaYolu = WordBelgesiOlustur(sIcerik, _
                 "Uygunluk_" & col("TescilNo") & "_" & Format(Date, "YYYYMMDD"))

    If sDosyaYolu <> "" Then
        MsgBox "Uygunluk Belgesi oluşturuldu:" & vbCrLf & sDosyaYolu, vbInformation, UYGULAMA_ADI
    End If
End Sub

'==============================================================================
' YARDIMCI FONKSİYONLAR (Private)
'==============================================================================

'------------------------------------------------------------------------------
' ŞABLON OKU
' ŞABLONLAR sayfasından belirtilen şablon adına göre metin döndürür.
'------------------------------------------------------------------------------
Private Function SablonOku(ByVal sSablonAdi As String) As String
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_SABLONLAR)
    If ws Is Nothing Then SablonOku = "": Exit Function

    Dim nSon As Long, i As Long
    nSon = SonSatir(ws, 1)
    For i = 1 To nSon
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(sSablonAdi) Then
            SablonOku = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    SablonOku = ""
End Function

'------------------------------------------------------------------------------
' DEĞİŞKENLERİ DOLDUR
' Şablon metnindeki <<DEĞIŞKEN>> yer tutucularını gerçek değerlerle değiştirir.
'------------------------------------------------------------------------------
Private Function DegiskenleriDoldur(ByVal sMetin As String, _
                                     ByVal degMap() As String) As String
    Dim s As String
    s = sMetin
    Dim i As Long
    For i = 0 To UBound(degMap, 1)
        If degMap(i, 0) <> "" Then
            s = Replace(s, degMap(i, 0), degMap(i, 1))
        End If
    Next i
    DegiskenleriDoldur = s
End Function

'------------------------------------------------------------------------------
' KURUM BİLGİLERİNİ DOLDUR
' degMap dizisinin 0-7 arası indekslerini AYARLAR sayfasından doldurur.
'------------------------------------------------------------------------------
Private Sub KurumBilgileriniDoldur(ByRef degMap() As String)
    degMap(0, 0) = "<<KURUM_ADI>>":      degMap(0, 1) = AyarOku("KURUM_ADI")
    degMap(1, 0) = "<<KURUM_BIRIM>>":    degMap(1, 1) = AyarOku("KURUM_BIRIM")
    degMap(2, 0) = "<<KURUM_ADRES>>":    degMap(2, 1) = AyarOku("KURUM_ADRES")
    degMap(3, 0) = "<<KURUM_TEL>>":      degMap(3, 1) = AyarOku("KURUM_TEL")
    degMap(4, 0) = "<<KURUM_EPOSTA>>":   degMap(4, 1) = AyarOku("KURUM_EPOSTA")
    degMap(5, 0) = "<<YETKILI_ADI>>":    degMap(5, 1) = AyarOku("YETKILI_ADI")
    degMap(6, 0) = "<<YETKILI_UNVAN>>":  degMap(6, 1) = AyarOku("YETKILI_UNVAN")
    degMap(7, 0) = "<<EVRAK_SAYISI>>":   degMap(7, 1) = YeniEvrakSayisi()
End Sub

'------------------------------------------------------------------------------
' YENİ EVRAK SAYISI ÜRET
' AYARLAR sayfasındaki sayacı artırır ve yeni sayıyı döndürür.
'------------------------------------------------------------------------------
Private Function YeniEvrakSayisi() As String
    Dim nMevcutSayac As Long
    Dim sMevcutSayac As String
    sMevcutSayac = AyarOku("EVRAK_SAYAC")
    If IsNumeric(sMevcutSayac) Then
        nMevcutSayac = CLng(sMevcutSayac) + 1
    Else
        nMevcutSayac = 1
    End If
    AyarYaz "EVRAK_SAYAC", CStr(nMevcutSayac)
    YeniEvrakSayisi = AyarOku("KURUM_BIRIM_KOD") & "/" & Year(Date) & "-" & _
                      Format(nMevcutSayac, "0000")
End Function

'------------------------------------------------------------------------------
' ASANSÖR ID SEÇ
' Kullanıcıdan asansör ID'si alır ve doğrular.
'------------------------------------------------------------------------------
Private Function AsansorIDSec() As Long
    Dim sID As String
    sID = Trim(InputBox("İşlem yapılacak Asansörün Kayıt ID'sini girin" & vbCrLf & _
                        "(Asansör ID'sini bilmiyorsanız önce ASANSÖRLER sayfasına bakın):", _
                        UYGULAMA_ADI & " - Asansör Seç", ""))
    If sID = "" Or Not IsNumeric(sID) Then
        AsansorIDSec = 0
        Exit Function
    End If
    AsansorIDSec = CLng(sID)
End Function

'------------------------------------------------------------------------------
' WORD BELGESİ OLUŞTUR
' Microsoft Word COM otomasyonu ile .docx belgesi oluşturur ve kaydeder.
' Kayıt yolu: Kayıt klasörü (AYARLAR'dan) veya Excel dosyasının bulunduğu yer.
' Döndürür: Kaydedilen dosyanın tam yolu (hata varsa boş string)
'------------------------------------------------------------------------------
Private Function WordBelgesiOlustur(ByVal sIcerik As String, ByVal sDosyaAdi As String) As String
    WordBelgesiOlustur = ""

    ' Kayıt klasörünü belirle
    Dim sKlasor As String
    sKlasor = AyarOku("KAYIT_KLASORU")
    If sKlasor = "" Then
        sKlasor = ThisWorkbook.Path & "\Yazılar"
    End If

    ' Klasör yoksa oluştur
    If Not (CreateObject("Scripting.FileSystemObject").FolderExists(sKlasor)) Then
        On Error Resume Next
        MkDir sKlasor
        On Error GoTo 0
    End If

    Dim sTamYol As String
    sTamYol = sKlasor & "\" & sDosyaAdi & ".docx"

    ' Word COM nesnesi oluştur (geç bağlama - Word kurulu olmak zorunda)
    Dim oWord As Object
    Dim oDoc As Object
    Dim oParagraf As Object

    On Error GoTo WordHata
    oWord = CreateObject("Word.Application")
    oWord.Visible = False

    Set oDoc = oWord.Documents.Add()

    ' Sayfa yapısı - A4 ve Türk resmî yazı kenar boşlukları
    With oDoc.PageSetup
        .PaperSize = 9          ' wdPaperA4
        .TopMargin = oWord.CentimetersToPoints(2.5)
        .BottomMargin = oWord.CentimetersToPoints(2.5)
        .LeftMargin = oWord.CentimetersToPoints(2.5)
        .RightMargin = oWord.CentimetersToPoints(2)
    End With

    ' Yazı tipini ve boyutunu ayarla
    With oDoc.Content.Font
        .Name = "Times New Roman"
        .Size = 12
    End With

    ' İçeriği satır satır yaz
    Dim satirlar() As String
    satirlar = Split(sIcerik, vbCrLf)
    Dim i As Long
    For i = 0 To UBound(satirlar)
        If i = 0 Then
            oDoc.Content.InsertAfter satirlar(i)
        Else
            oDoc.Content.InsertAfter vbCr & satirlar(i)
        End If
    Next i

    ' Belgeyi kaydet ve kapat
    oDoc.SaveAs2 FileName:=sTamYol, FileFormat:=12  ' 12 = docx
    oDoc.Close SaveChanges:=False
    oWord.Quit

    Set oDoc = Nothing
    Set oWord = Nothing

    WordBelgesiOlustur = sTamYol

    ' Word'ü aç (kullanıcı isteğine göre)
    If Onayla("Belge oluşturuldu. Şimdi Word'de açılsın mı?" & vbCrLf & sTamYol) Then
        Shell "explorer.exe """ & sTamYol & """", vbNormalFocus
    End If

    Exit Function

WordHata:
    ' Word kurulu değilse metin dosyası olarak kaydet
    On Error Resume Next
    If Not oDoc Is Nothing Then oDoc.Close False
    If Not oWord Is Nothing Then oWord.Quit
    On Error GoTo 0

    ' .txt formatında kaydet
    Dim sTxtYol As String
    sTxtYol = sKlasor & "\" & sDosyaAdi & ".txt"
    Dim nDosya As Integer
    nDosya = FreeFile
    Open sTxtYol For Output As #nDosya
    Print #nDosya, sIcerik
    Close #nDosya

    MsgBox "Word bulunamadı veya hata oluştu. Yazı metin dosyası olarak kaydedildi:" & _
           vbCrLf & sTxtYol, vbInformation, UYGULAMA_ADI
    WordBelgesiOlustur = sTxtYol
End Function

'==============================================================================
' GÜVENSIZ ASANSÖR İDARİ YAPTIRIM ÜST YAZISI
'==============================================================================

'------------------------------------------------------------------------------
' GÜVENSIZ ASANSÖR İDARİ YAPTIRIM ÜST YAZISI ÜRETİMİ
' "Güvensiz" (ciddi riskli) asansörler için idari para cezası üst yazısı.
' 7223 sayılı Ürün Güvenliği ve Teknik Düzenlemeler Kanunu kapsamında.
' Tutanak .docx dosyasından otomatik bilgi yükleme destekler.
'------------------------------------------------------------------------------
Public Sub GuvensizAsansorYaptirimUret(ByVal nAsansorID As Long)
    Dim col As Collection
    Set col = AsansorGetir(nAsansorID)
    If col.Count = 0 Then
        MsgBox "Asansör bulunamadı!", vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If

    Dim sSablon As String
    sSablon = SablonOku("IDARI_YAPTIRIM_USTYAZI")
    If sSablon = "" Then
        MsgBox "İdari Yaptırım Üst Yazısı şablonu ŞABLONLAR sayfasında bulunamadı!", _
               vbCritical, UYGULAMA_ADI
        Exit Sub
    End If

    ' ── Tutanaktan otomatik bilgi yükleme ──────────────────────────────────
    Dim colTutanak As Collection
    Set colTutanak = Nothing

    If Onayla("Tutanak dosyasından (.docx) otomatik bilgi yüklemek ister misiniz?") Then
        Dim sTutanakDosyasi As String
        sTutanakDosyasi = DosyaSec("Word Belgesi (*.docx)", "*.docx")
        If sTutanakDosyasi <> "" Then
            Set colTutanak = TutanakOku(sTutanakDosyasi)
            If colTutanak Is Nothing Then
                MsgBox "Tutanak dosyası okunamadı. Tüm alanlar manuel girilecek.", _
                       vbInformation, UYGULAMA_ADI
            End If
        End If
    End If

    ' ── Ön değerleri belirle (tutanaktan veya envanterden) ─────────────────
    Dim sFirmaAdi As String, sVergiNo As String, sFirmaAdresi As String
    Dim sAsansorAdres As String, sAsansorTip As String
    Dim sTespit As String

    sFirmaAdi    = TutanakDeger(colTutanak, "FIRMA_ADI")
    sVergiNo     = TutanakDeger(colTutanak, "VERGI_NO")
    sFirmaAdresi = TutanakDeger(colTutanak, "FIRMA_ADRESI")

    ' Tespit tarih/sayı: tutanaktan veya boş
    Dim sTespit_tar As String, sTespit_sayi As String
    sTespit_tar  = TutanakDeger(colTutanak, "TARIH")
    sTespit_sayi = TutanakDeger(colTutanak, "TUTANAK_NO")
    If sTespit_tar <> "" And sTespit_sayi <> "" Then
        sTespit = sTespit_tar & " - " & sTespit_sayi
    End If

    ' Asansör adresi: tutanaktan yoksa envanterden
    sAsansorAdres = TutanakDeger(colTutanak, "ASANSOR_ADRES")
    If sAsansorAdres = "" Then sAsansorAdres = col("Adres") & " " & col("Ilce")

    ' Asansör tipi: tutanaktan yoksa envanterden
    sAsansorTip = TutanakDeger(colTutanak, "ASANSOR_TIP")
    If sAsansorTip = "" Then sAsansorTip = col("Tip")

    ' ── Kullanıcı girişleri ─────────────────────────────────────────────────
    sFirmaAdi = Trim(InputBox( _
        "Cezalandırılan firma / kişi adı:", UYGULAMA_ADI, sFirmaAdi))
    If sFirmaAdi = "" Then
        MsgBox "Firma adı boş bırakılamaz.", vbExclamation, UYGULAMA_ADI
        Exit Sub
    End If

    Dim sVergiDairesi As String
    sVergiDairesi = Trim(InputBox( _
        "Vergi dairesi adı:", UYGULAMA_ADI, ""))
    sVergiNo = Trim(InputBox( _
        "Vergi numarası:", UYGULAMA_ADI, sVergiNo))
    Dim sTCKimlik As String
    sTCKimlik = Trim(InputBox( _
        "T.C. kimlik no (tüzel kişi ise '-' girin):", UYGULAMA_ADI, "-"))
    sFirmaAdresi = Trim(InputBox( _
        "Firma adresi:", UYGULAMA_ADI, sFirmaAdresi))
    sTespit = Trim(InputBox( _
        "Tespit raporu/tutanağı tarih ve sayısı:" & vbCrLf & _
        "(örn: 25.07.2024 - 20-12DD69D5)", UYGULAMA_ADI, sTespit))
    sAsansorAdres = Trim(InputBox( _
        "Asansörün bulunduğu adres:", UYGULAMA_ADI, sAsansorAdres))
    sAsansorTip = Trim(InputBox( _
        "Asansör tipi:", UYGULAMA_ADI, sAsansorTip))

    Dim sPGDKurulus As String
    sPGDKurulus = Trim(InputBox( _
        "PGD testi yapan kuruluş adı:", UYGULAMA_ADI, ""))
    Dim sPGDTar As String
    sPGDTar = Trim(InputBox( _
        "PGD test raporu tarihi (GG/AA/YYYY):", UYGULAMA_ADI, ""))
    Dim sPGDYaziSayi As String
    sPGDYaziSayi = Trim(InputBox( _
        "PGD kuruluşunun yazı sayısı:", UYGULAMA_ADI, ""))
    Dim sPGDBelgeNo As String
    sPGDBelgeNo = Trim(InputBox( _
        "PGD test raporu numarası:", UYGULAMA_ADI, ""))

    Dim sBakYaziTar As String
    sBakYaziTar = Trim(InputBox( _
        "Bakanlık kılavuz yazısı tarihi:", UYGULAMA_ADI, "24.03.2022"))
    Dim sBakYaziSayi As String
    sBakYaziSayi = Trim(InputBox( _
        "Bakanlık yazısının sayısı:", UYGULAMA_ADI, "3473190"))

    Dim sRiskTar As String
    sRiskTar = Trim(InputBox( _
        "Risk değerlendirme tablosu tarihi:", UYGULAMA_ADI, TarihStr(Date)))
    Dim sRiskSira As String
    sRiskSira = Trim(InputBox( _
        "Risk değerlendirme sıra numarası (örn: 2024/2):", _
        UYGULAMA_ADI, Year(Date) & "/1"))

    Dim sBildirimTar As String
    sBildirimTar = Trim(InputBox( _
        "Firmaya önceki bildirim yazısı tarihi:", UYGULAMA_ADI, ""))
    Dim sBildirimSayi As String
    sBildirimSayi = Trim(InputBox( _
        "Önceki bildirim yazısının sayısı:", UYGULAMA_ADI, ""))

    Dim sCezaTutari As String
    sCezaTutari = Trim(InputBox( _
        "İdari para cezası tutarı rakamla (örn: 115.226,00):", UYGULAMA_ADI, ""))
    Dim sCezaTutariYazi As String
    sCezaTutariYazi = Trim(InputBox( _
        "İdari para cezası tutarı yazıyla:" & vbCrLf & _
        "(örn: Yüz OnBeşBinİkiYüzYirmialtı)", UYGULAMA_ADI, ""))

    ' ── degMap doldur ────────────────────────────────────────────────────────
    Dim degMap() As String
    ReDim degMap(0 To 28, 0 To 1)
    KurumBilgileriniDoldur degMap
    degMap(8, 0) = "<<FIRMA_ADI>>":           degMap(8, 1) = sFirmaAdi
    degMap(9, 0) = "<<VERGI_DAIRESI>>":       degMap(9, 1) = sVergiDairesi
    degMap(10, 0) = "<<VERGI_NO>>":           degMap(10, 1) = sVergiNo
    degMap(11, 0) = "<<TC_KIMLIK>>":          degMap(11, 1) = sTCKimlik
    degMap(12, 0) = "<<FIRMA_ADRESI>>":       degMap(12, 1) = sFirmaAdresi
    degMap(13, 0) = "<<TESPIT_TAR_SAYI>>":    degMap(13, 1) = sTespit
    degMap(14, 0) = "<<ASANSOR_ADRES>>":      degMap(14, 1) = sAsansorAdres
    degMap(15, 0) = "<<ASANSOR_TIP>>":        degMap(15, 1) = sAsansorTip
    degMap(16, 0) = "<<PGD_KURULUS>>":        degMap(16, 1) = sPGDKurulus
    degMap(17, 0) = "<<PGD_TAR>>":            degMap(17, 1) = sPGDTar
    degMap(18, 0) = "<<PGD_YAZI_SAYI>>":      degMap(18, 1) = sPGDYaziSayi
    degMap(19, 0) = "<<PGD_BELGE_NO>>":       degMap(19, 1) = sPGDBelgeNo
    degMap(20, 0) = "<<BAKANLIK_YAZI_TAR>>":  degMap(20, 1) = sBakYaziTar
    degMap(21, 0) = "<<BAKANLIK_YAZI_SAYI>>": degMap(21, 1) = sBakYaziSayi
    degMap(22, 0) = "<<RISK_DEG_TAR>>":       degMap(22, 1) = sRiskTar
    degMap(23, 0) = "<<RISK_DEG_SIRA>>":      degMap(23, 1) = sRiskSira
    degMap(24, 0) = "<<BILDIRIM_TAR>>":       degMap(24, 1) = sBildirimTar
    degMap(25, 0) = "<<BILDIRIM_SAYI>>":      degMap(25, 1) = sBildirimSayi
    degMap(26, 0) = "<<CEZA_TUTARI>>":        degMap(26, 1) = sCezaTutari
    degMap(27, 0) = "<<CEZA_TUTARI_YAZI>>":   degMap(27, 1) = sCezaTutariYazi
    degMap(28, 0) = "<<TARIH>>":              degMap(28, 1) = TarihStr(Date)

    Dim sIcerik As String
    sIcerik = DegiskenleriDoldur(sSablon, degMap)

    Dim sDosyaYolu As String
    sDosyaYolu = WordBelgesiOlustur(sIcerik, _
                 "GuvensizAsansor_Yaptirim_" & col("TescilNo") & "_" & Format(Date, "YYYYMMDD"))

    If sDosyaYolu <> "" Then
        MsgBox "İdari Yaptırım Üst Yazısı oluşturuldu:" & vbCrLf & sDosyaYolu, _
               vbInformation, UYGULAMA_ADI
    End If
End Sub

'------------------------------------------------------------------------------
' TUTANAK DOSYASINDAN BİLGİ OKU
' Sanayi Bakanlığı PGD tutanak .docx dosyasını okur ve
' tablo hücrelerini etiket→değer çiftleri olarak döndürür.
' Döndürür: Collection (key=anahtar, value=değer) veya Nothing hata durumunda.
'------------------------------------------------------------------------------
Public Function TutanakOku(ByVal sDocxYol As String) As Collection
    Set TutanakOku = Nothing
    If sDocxYol = "" Then Exit Function

    ' Geçici dosya yolları
    Dim sTmpZip As String, sTmpXml As String
    sTmpZip = Environ("TEMP") & "\pgd_tutanak_tmp.zip"
    sTmpXml = Environ("TEMP") & "\pgd_tutanak_doc.xml"

    On Error GoTo TutanakHata

    ' .docx'ı .zip olarak kopyala (OOXML formatı ZIP tabanlıdır)
    FileCopy sDocxYol, sTmpZip

    ' PowerShell ile ZIP içindeki word/document.xml'i oku
    Dim oWsh As Object
    Set oWsh = CreateObject("WScript.Shell")
    Dim sPSCmd As String
    sPSCmd = "Add-Type -AssemblyName System.IO.Compression.FileSystem; " & _
             "$zip = [System.IO.Compression.ZipFile]::OpenRead('" & sTmpZip & "'); " & _
             "$entry = $zip.Entries | Where-Object {$_.FullName -eq 'word/document.xml'}; " & _
             "$stream = $entry.Open(); " & _
             "$reader = New-Object System.IO.StreamReader($stream); " & _
             "$content = $reader.ReadToEnd(); " & _
             "$reader.Close(); $stream.Close(); $zip.Dispose(); " & _
             "[System.IO.File]::WriteAllText('" & sTmpXml & "', " & _
             "$content, [System.Text.Encoding]::UTF8)"
    oWsh.Run "powershell -NoProfile -Command """ & sPSCmd & """", 0, True
    Set oWsh = Nothing

    ' XML'i parse et
    Dim oXML As Object
    Set oXML = CreateObject("MSXML2.DOMDocument.6.0")
    oXML.async = False
    If Not oXML.Load(sTmpXml) Then GoTo TutanakHata

    oXML.setProperty "SelectionNamespaces", _
        "xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'"

    ' Tüm tablo hücrelerinin düz metnini bir diziye al
    Dim oTcNodes As Object
    Set oTcNodes = oXML.selectNodes("//w:tc")
    Dim n As Long
    n = oTcNodes.Length
    If n = 0 Then GoTo TutanakHata

    Dim aCells() As String
    ReDim aCells(0 To n - 1)
    Dim i As Long, j As Long
    For i = 0 To n - 1
        Dim oTNodes As Object
        Set oTNodes = oTcNodes(i).selectNodes(".//w:t")
        Dim sCellText As String
        sCellText = ""
        For j = 0 To oTNodes.Length - 1
            sCellText = sCellText & oTNodes(j).Text
        Next j
        aCells(i) = Trim(sCellText)
    Next i

    ' Etiket→değer eşlemesi
    ' Başlık satırı grubu: "Tarih", "Tutanak No" etiketleri için offset=3
    ' (3 etiket hücresi ardından 3 değer hücresi gelir)
    ' Detay bölümü: bitişik hücre çiftleri için offset=1
    Dim col As New Collection
    SozcukEkle col, aCells, "Tarih", 3, "TARIH"
    SozcukEkle col, aCells, "Tutanak No", 3, "TUTANAK_NO"
    SozcukEkle col, aCells, "Vergi No/T.C. No", 1, "VERGI_NO"
    SozcukEkle col, aCells, "Ticaret Ünvanı/ Adı Soyadı", 1, "FIRMA_ADI"
    SozcukEkle col, aCells, "Adresi", 1, "FIRMA_ADRESI"
    SozcukEkle col, aCells, "Montaj Adresi", 1, "ASANSOR_ADRES"
    SozcukEkle col, aCells, "Ürün Adı", 1, "ASANSOR_TIP"

    ' Ürün adını temizle: "ASANSÖR > ASANSÖR > ELEKTRİKLİ ASANSÖR" → son parça
    On Error Resume Next
    Dim sUrunAdi As String
    sUrunAdi = col("ASANSOR_TIP")
    If InStr(sUrunAdi, ">") > 0 Then
        Dim aUrunParcalar() As String
        aUrunParcalar = Split(sUrunAdi, ">")
        col.Remove "ASANSOR_TIP"
        col.Add Trim(aUrunParcalar(UBound(aUrunParcalar))), "ASANSOR_TIP"
    End If
    On Error GoTo TutanakHata

    ' Geçici dosyaları temizle
    On Error Resume Next
    Kill sTmpZip
    Kill sTmpXml
    On Error GoTo 0

    Set TutanakOku = col
    Exit Function

TutanakHata:
    On Error Resume Next
    Kill sTmpZip
    Kill sTmpXml
    Set TutanakOku = Nothing
End Function

'------------------------------------------------------------------------------
' TUTANAK KOLEKSİYONUNDAN DEĞER AL
' Collection'da anahtar yoksa boş string döndürür (hata bastırır).
'------------------------------------------------------------------------------
Private Function TutanakDeger(ByVal col As Collection, ByVal sAnahtar As String) As String
    If col Is Nothing Then TutanakDeger = "": Exit Function
    On Error Resume Next
    TutanakDeger = col(sAnahtar)
    If Err.Number <> 0 Then TutanakDeger = ""
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' SÖZCÜK EKLE (Tutanak yardımcısı)
' Hücre dizisinde etiket arar ve (etiket + offset) konumundaki değeri
' Collection'a ekler. Etiket bulunamazsa veya değer boşsa işlem yapılmaz.
'------------------------------------------------------------------------------
Private Sub SozcukEkle(ByRef col As Collection, ByVal aCells() As String, _
                        ByVal sEtiket As String, ByVal nOffset As Long, _
                        ByVal sAnahtar As String)
    Dim i As Long
    For i = 0 To UBound(aCells) - nOffset
        If aCells(i) = sEtiket Then
            Dim sDeger As String
            sDeger = aCells(i + nOffset)
            If sDeger <> "" Then
                On Error Resume Next
                col.Remove sAnahtar
                On Error GoTo 0
                col.Add sDeger, sAnahtar
            End If
            Exit Sub
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' DOSYA SEÇ
' Excel FileDialog ile kullanıcıya dosya seçtirme iletişim kutusu açar.
' Döndürür: Seçilen dosyanın tam yolu, iptal edilirse boş string.
'------------------------------------------------------------------------------
Private Function DosyaSec(ByVal sAciklama As String, ByVal sFiltre As String) As String
    DosyaSec = ""
    On Error GoTo DosyaSecHata
    Dim fd As Object
    Set fd = Application.FileDialog(3)  ' msoFileDialogFilePicker
    With fd
        .Title = "Tutanak Dosyası Seç"
        .Filters.Clear
        .Filters.Add sAciklama, sFiltre
        .AllowMultiSelect = False
        If .Show = -1 Then DosyaSec = .SelectedItems(1)
    End With
    Exit Function
DosyaSecHata:
    DosyaSec = ""
End Function
