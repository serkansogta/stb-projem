Attribute VB_Name = "ThisWorkbook"
'==============================================================================
' ASANSÖR PGD TAKİP SİSTEMİ - İŞ KİTABI OLAYLARI
' Modül : ThisWorkbook
' Açıklama: Excel iş kitabı açılış/kapanış olaylarını yönetir.
'           Bu modül Excel'de ThisWorkbook nesnesine kopyalanmalıdır.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' İş kitabı açıldığında GİRİŞ sayfasını güncelle
'------------------------------------------------------------------------------
Private Sub Workbook_Open()
    On Error Resume Next
    Application.ScreenUpdating = False

    ' Asansörlerin PGD renk durumlarını yenile (sessiz mod)
    Dim ws As Worksheet
    Set ws = SayfaBul(SHT_ASANSORLER)
    If Not ws Is Nothing Then
        SayfaKorumaKapat ws
        Dim nSon As Long, i As Long
        nSon = SonSatir(ws, 1)
        For i = 2 To nSon
            If ws.Cells(i, ASN_ID).Value <> "" Then
                Dim dSon As Variant
                dSon = ws.Cells(i, ASN_SON_PGD_TAR).Value
                If IsDate(dSon) And Not IsEmpty(dSon) Then
                    ws.Cells(i, ASN_SON_PGD_TAR2).Value = CDate(dSon) + PGD_SURE_GUN
                    ws.Cells(i, ASN_SON_PGD_TAR2).NumberFormat = "DD.MM.YYYY"
                End If
                AsansorleriSatirGuncelle ws, i
            End If
        Next i
        SayfaKorumaAc ws
    End If

    ' Dashboard'u güncelle
    GirisSayfasiGuncelle

    ' GİRİŞ sayfasına git
    On Error Resume Next
    ThisWorkbook.Worksheets(SHT_GIRIS).Activate
    On Error GoTo 0

    Application.ScreenUpdating = True

    ' Kritik uyarı mesajı
    Dim nGecikmiş As Long
    nGecikmiş = GecikmisMiktar()
    If nGecikmiş > 0 Then
        MsgBox "UYARI: " & nGecikmiş & " asansörün PGD süresi dolmuş!" & vbCrLf & _
               "GİRİŞ sayfasından Takvim Güncelle veya Durum Raporu alabilirsiniz.", _
               vbExclamation, UYGULAMA_ADI
    End If
End Sub

'------------------------------------------------------------------------------
' İş kitabı kaydedilmeden önce
'------------------------------------------------------------------------------
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Kayıt öncesi tarihi güncelle (isteğe bağlı)
End Sub
