# Asansör PGD Takip Sistemi — Kurulum Kılavuzu

## Genel Bakış

Bu sistem, belediyelerin ve yetkili denetim birimlerinin asansör Periyodik Güvenlik
Denetimi (PGD) süreçlerini yönetmesi için geliştirilmiş Excel/VBA tabanlı bir uygulamadır.

### Özellikler

| Özellik | Açıklama |
|---|---|
| Asansör Envanteri | Bina, malik, tescil no, tip, kapasite bilgileri |
| PGD Takip | Son PGD tarihi, sonraki PGD tarihi, gecikme durumu |
| Renk Kodlama | Kırmızı=Gecikmiş, Turuncu=Kritik, Sarı=Yaklaşıyor, Yeşil=Uygun |
| Matbu Yazılar | 4 farklı resmi yazı türü otomatik üretimi |
| Takvim | Tüm asansörlerin PGD durumu listelenmiş takvim görünümü |
| Raporlar | İlçe bazlı istatistik, gecikmiş liste, özet dashboard |
| Toplu Bildirim | Tek tıkla tüm gecikmiş asansörlere yazı üretimi |

---

## Kurulum Adımları

### 1. Excel Dosyasını Hazırlayın

1. Microsoft Excel'i açın
2. Yeni boş bir çalışma kitabı oluşturun
3. **Makrolu çalışma kitabı** olarak kaydedin: `Asansör_PGD_Takip.xlsm`

### 2. VBA Geliştirici Sekmesini Etkinleştirin

`Dosya → Seçenekler → Şeridi Özelleştir → Geliştirici (işaretleyin) → Tamam`

### 3. VBA Modüllerini İçe Aktarın

`Alt + F11` tuşlarına basarak VBA Düzenleyicisini açın.

**Her `.bas` dosyası için:**
1. VBA Düzenleyicide `Dosya → Dosya İçe Aktar` (veya `Ctrl+M`)
2. `src/` klasöründeki dosyaları sırayla seçin:
   - `modSabitler.bas`
   - `modYardimci.bas`
   - `modEnvanter.bas`
   - `modPGD.bas`
   - `modMatbuYazi.bas`
   - `modRaporlama.bas`
   - `modKurulum.bas`

**ThisWorkbook için:**
1. VBA Düzenleyicide sol panelde `ThisWorkbook`'a çift tıklayın
2. `src/ThisWorkbook.bas` dosyasını açın, içeriği kopyalayın ve yapıştırın

### 4. Kurulum Makrosunu Çalıştırın

`Alt + F8 → KurulumYap → Çalıştır`

Bu işlem tüm sayfaları, başlıkları ve şablonları otomatik olarak oluşturacaktır.

### 5. Kurum Bilgilerini Doldurun

`AYARLAR` sayfasını açın ve aşağıdaki alanları kendi kurumunuza göre doldurun:

| Anahtar | Açıklama | Örnek |
|---|---|---|
| `KURUM_ADI` | Kurumun tam adı | Kadıköy Belediyesi |
| `KURUM_BIRIM` | Birim adı | Ruhsat ve Denetim Müdürlüğü |
| `KURUM_BIRIM_KOD` | Kısa kod (evrak sayısında) | RDM |
| `KURUM_ADRES` | Kurum adresi | Mühürdar Cad. No:20 Kadıköy/İstanbul |
| `KURUM_TEL` | Telefon | 0 (216) XXX XX XX |
| `KURUM_EPOSTA` | E-posta | rdm@kadikoy.bel.tr |
| `YETKILI_ADI` | İmzalayacak yetkili | Ahmet Yılmaz |
| `YETKILI_UNVAN` | Unvanı | Müdür |
| `KAYIT_KLASORU` | Yazıların kaydedileceği yer | C:\PGD_Yazılar (boş=otomatik) |

---

## Kullanım

### Yeni Asansör Eklemek

`GİRİŞ sayfası → Yeni Asansör Ekle` düğmesi veya `Alt+F8 → AsansorEkle`

### PGD Kaydı Girmek

`GİRİŞ sayfası → PGD Kaydı Ekle` düğmesi → Asansör ID'sini girin

### Matbu Yazı Üretmek

`GİRİŞ sayfası → Matbu Yazı Üret` → Yazı türünü seçin → Asansör ID'sini girin

Yazılar Word kuruluysa `.docx`, kurulu değilse `.txt` olarak kaydedilir.

### PGD Takvimini Güncellemek

`GİRİŞ sayfası → Takvim Güncelle`

Takvim sayfası renk kodlu şekilde sıralanır:
- **Kırmızı** — PGD süresi dolmuş (gecikmiş)
- **Turuncu** — 0–15 gün kaldı (kritik)
- **Sarı** — 16–60 gün kaldı (yaklaşıyor)
- **Yeşil** — 60 günden fazla kaldı (uygun)

### Toplu Bildirim

`GİRİŞ sayfası → Toplu PGD Bildirimi` — Gecikmiş/yaklaşan tüm asansörler için
PGD Bildirimi yazısı otomatik üretilir.

---

## Proje Dosya Yapısı

```
stb-projem/
├── src/
│   ├── modSabitler.bas      # Sabitler, sütun numaraları, renkler
│   ├── modYardimci.bas      # Genel amaçlı yardımcı fonksiyonlar
│   ├── modEnvanter.bas      # Asansör ekleme/düzenleme/silme
│   ├── modPGD.bas           # PGD kayıt girişi, takvim yönetimi
│   ├── modMatbuYazi.bas     # Word belgesi otomasyonu (4 yazı türü)
│   ├── modRaporlama.bas     # Dashboard, durum raporu, istatistikler
│   ├── modKurulum.bas       # İlk kurulum, sayfa oluşturma, şablonlar
│   └── ThisWorkbook.bas     # Workbook_Open olayı
└── sablonlar/
    ├── pgd_bildirimi.txt    # PGD Bildirimi şablon referansı
    ├── ihtar_yazisi.txt     # İhtar Yazısı şablon referansı
    ├── muhürleme_karari.txt # Mühürleme Kararı şablon referansı
    └── uygunluk_belgesi.txt # Uygunluk Belgesi şablon referansı
```

---

## Şablon Değişkenleri

Şablonlardaki `<<DEĞİŞKEN>>` yer tutucuları, yazı üretilirken otomatik doldurulur.
`ŞABLONLAR` sayfasındaki metinleri düzenleyerek yazıları özelleştirebilirsiniz.

---

## Yasal Dayanak

- 2014/33/EU Asansör Yönetmeliği (29.06.2016 tarih ve 29757 sayılı RG)
- 6331 sayılı İş Sağlığı ve Güvenliği Kanunu
- Asansör İşletme, Bakım ve Periyodik Kontrol Yönetmeliği
