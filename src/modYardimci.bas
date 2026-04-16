Attribute VB_Name = "modYardimci"
'==============================================================================
' ASANSÖR PGD TAKİP SİSTEMİ - YARDIMCI FONKSİYONLAR
' Modül : modYardimci
' Açıklama: Tüm modüller tarafından kullanılan genel amaçlı yardımcı
'           fonksiyonlar ve prosedürler.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' HATA YÖNETİMİ YARDIMCISI
' Hata mesajını standart formatta kullanıcıya gösterir.
'------------------------------------------------------------------------------
Public Sub HataGoster(ByVal sProsedur As String, ByVal oErr As ErrObject)
    MsgBox "Hata oluştu!" & vbCrLf & vbCrLf & _
           "Prosedür : " & sProsedur & vbCrLf & _
           "Hata No  : " & oErr.Number & vbCrLf & _
           "Açıklama : " & oErr.Description, _
           vbCritical, "Asansör PGD Sistemi - Hata"
End Sub

'------------------------------------------------------------------------------
' SAYFA BULMA
' Adına göre sayfa nesnesi döndürür. Sayfa bulunamazsa Nothing döner.
'------------------------------------------------------------------------------
Public Function SayfaBul(ByVal sSayfaAdi As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sSayfaAdi)
    On Error GoTo 0
    Set SayfaBul = ws
End Function

'------------------------------------------------------------------------------
' SAYFADA SON SATIRI BUL
' Verilen sütundaki son dolu satır numarasını döndürür.
'------------------------------------------------------------------------------
Public Function SonSatir(ByVal ws As Worksheet, Optional ByVal nSutun As Integer = 1) As Long
    SonSatir = ws.Cells(ws.Rows.Count, nSutun).End(xlUp).Row
End Function

'------------------------------------------------------------------------------
' YENİ ID ÜRET
' Verilen sayfadaki A sütununda mevcut en büyük ID'yi bulup 1 artırır.
'------------------------------------------------------------------------------
Public Function YeniID(ByVal ws As Worksheet) As Long
    Dim nSon As Long
    Dim nMaks As Long
    Dim i As Long
    nSon = SonSatir(ws, 1)
    nMaks = 0
    For i = 2 To nSon
        If IsNumeric(ws.Cells(i, 1).Value) Then
            If CLng(ws.Cells(i, 1).Value) > nMaks Then
                nMaks = CLng(ws.Cells(i, 1).Value)
            End If
        End If
    Next i
    YeniID = nMaks + 1
End Function

'------------------------------------------------------------------------------
' TARİH FORMATLAMA
' Tarihi "DD.MM.YYYY" formatında string olarak döndürür.
'------------------------------------------------------------------------------
Public Function TarihStr(ByVal dTarih As Variant) As String
    If IsDate(dTarih) And Not IsEmpty(dTarih) Then
        TarihStr = Format(CDate(dTarih), "DD.MM.YYYY")
    Else
        TarihStr = ""
    End If
End Function

'------------------------------------------------------------------------------
' PGD DURUM RENGİ
' Son PGD tarihine göre hücre arka plan rengini döndürür.
'   Kırmızı : Gecikmiş veya tarih yok
'   Turuncu : Kritik eşiğin altında (<=15 gün)
'   Sarı    : Uyarı eşiğinin altında (<=60 gün)
'   Yeşil   : Uygun (>60 gün)
'------------------------------------------------------------------------------
Public Function PGDDurumRengi(ByVal dSonrakiTarih As Variant) As Long
    If IsEmpty(dSonrakiTarih) Or Not IsDate(dSonrakiTarih) Then
        PGDDurumRengi = RENK_TEHLIKE
        Exit Function
    End If
    Dim nKalanGun As Long
    nKalanGun = CDate(dSonrakiTarih) - Date
    If nKalanGun < 0 Then
        PGDDurumRengi = RENK_TEHLIKE       ' Kırmızı - gecikmiş
    ElseIf nKalanGun <= PGD_KRITIK_ESIGI Then
        PGDDurumRengi = RENK_TEHLIKE       ' Kırmızı - kritik
    ElseIf nKalanGun <= PGD_UYARI_ESIGI Then
        PGDDurumRengi = RENK_UYARI        ' Turuncu - uyarı
    Else
        PGDDurumRengi = RENK_UYGUN        ' Yeşil - uygun
    End If
End Function

'------------------------------------------------------------------------------
' PGD DURUM METNİ
' Kalan gün sayısına göre açıklayıcı metin döndürür.
'------------------------------------------------------------------------------
Public Function PGDDurumMetni(ByVal dSonrakiTarih As Variant) As String
    If IsEmpty(dSonrakiTarih) Or Not IsDate(dSonrakiTarih) Then
        PGDDurumMetni = "PGD tarihi girilmemiş"
        Exit Function
    End If
    Dim nKalanGun As Long
    nKalanGun = CDate(dSonrakiTarih) - Date
    If nKalanGun < 0 Then
        PGDDurumMetni = Abs(nKalanGun) & " gün gecikmiş"
    ElseIf nKalanGun = 0 Then
        PGDDurumMetni = "Bugün son gün!"
    ElseIf nKalanGun <= PGD_KRITIK_ESIGI Then
        PGDDurumMetni = nKalanGun & " gün kaldı (KRİTİK)"
    ElseIf nKalanGun <= PGD_UYARI_ESIGI Then
        PGDDurumMetni = nKalanGun & " gün kaldı (Yaklaşıyor)"
    Else
        PGDDurumMetni = nKalanGun & " gün kaldı"
    End If
End Function

'------------------------------------------------------------------------------
' ONAY KUTUSU
' Evet/Hayır sorusu sorar, True/False döndürür.
'------------------------------------------------------------------------------
Public Function Onayla(ByVal sSoru As String) As Boolean
    Onayla = (MsgBox(sSoru, vbQuestion + vbYesNo, UYGULAMA_ADI) = vbYes)
End Function

'------------------------------------------------------------------------------
' SÜTUN HARFİ
' Sütun numarasından Excel sütun harfini döndürür (1=A, 26=Z, 27=AA ...)
'------------------------------------------------------------------------------
Public Function SutunHarfi(ByVal nSutun As Integer) As String
    Dim sHarf As String
    sHarf = ""
    Do While nSutun > 0
        Dim nKalan As Integer
        nKalan = (nSutun - 1) Mod 26
        sHarf = Chr(65 + nKalan) & sHarf
        nSutun = (nSutun - 1) \ 26
    Loop
    SutunHarfi = sHarf
End Function

'------------------------------------------------------------------------------
' HÜCRE STILI UYGULA - BAŞLIK SATIRI
'------------------------------------------------------------------------------
Public Sub BaslikStiliUygula(ByVal rng As Range)
    With rng
        .Interior.Color = RENK_BASLIK
        .Font.Color = RENK_BASLIK_YAZ
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 35
    End With
End Sub

'------------------------------------------------------------------------------
' HÜCRE STILI UYGULA - VERİ SATIRI
'------------------------------------------------------------------------------
Public Sub VeriStiliUygula(ByVal rng As Range, ByVal bCift As Boolean)
    With rng
        If bCift Then
            .Interior.Color = RENK_ZEBRA2
        Else
            .Interior.Color = RENK_ZEBRA1
        End If
        .Font.Color = RGB(50, 50, 50)
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With
End Sub

'------------------------------------------------------------------------------
' SAYFA KORUMA AÇ/KAPAT
'------------------------------------------------------------------------------
Public Sub SayfaKorumaKapat(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:="pgd2024"
    On Error GoTo 0
End Sub

Public Sub SayfaKorumaAc(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:="pgd2024", _
               DrawingObjects:=True, _
               Contents:=True, _
               AllowSorting:=True, _
               AllowFiltering:=True
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' AYAR DEĞER OKU
' AYARLAR sayfasından anahtar-değer çifti okur.
'------------------------------------------------------------------------------
Public Function AyarOku(ByVal sAnahtar As String) As String
    Dim ws As Worksheet
    Dim i As Long
    Dim nSon As Long
    AyarOku = ""
    Set ws = SayfaBul(SHT_AYARLAR)
    If ws Is Nothing Then Exit Function
    nSon = SonSatir(ws, 1)
    For i = 2 To nSon
        If Trim(ws.Cells(i, 1).Value) = sAnahtar Then
            AyarOku = Trim(ws.Cells(i, 2).Value)
            Exit For
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' AYAR DEĞER YAZ
' AYARLAR sayfasına anahtar-değer çifti yazar.
'------------------------------------------------------------------------------
Public Sub AyarYaz(ByVal sAnahtar As String, ByVal sDeger As String)
    Dim ws As Worksheet
    Dim i As Long
    Dim nSon As Long
    Set ws = SayfaBul(SHT_AYARLAR)
    If ws Is Nothing Then Exit Sub
    nSon = SonSatir(ws, 1)
    For i = 2 To nSon
        If Trim(ws.Cells(i, 1).Value) = sAnahtar Then
            ws.Cells(i, 2).Value = sDeger
            Exit Sub
        End If
    Next i
    ' Bulunamadıysa yeni satıra ekle
    ws.Cells(nSon + 1, 1).Value = sAnahtar
    ws.Cells(nSon + 1, 2).Value = sDeger
End Sub

'------------------------------------------------------------------------------
' SAYI YUVARLAMA - Para/miktar formatı için
'------------------------------------------------------------------------------
Public Function SayiFormat(ByVal dSayi As Double) As String
    SayiFormat = Format(dSayi, "#,##0")
End Function
