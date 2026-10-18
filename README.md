# Aydem · Günlük Üretim ve Ciro Sitesi

Her sabah 07:30'da GitHub kendiliğinden şunları yapar:

1. EPİAŞ Şeffaflık Platformu'ndan Aydem santrallerinin yeni günlerini çeker
2. Ciroyu hesaplar (YEKDEM'li santral: üretim × YEKDEM fiyatı × kur, diğerleri: saatlik üretim × saatlik PTF)
3. Siteyi günceller: grafikler, ortalama tablosu, santral tablosu, Excel indirme bağlantısı

Sen hiçbir şey çalıştırmazsın. Site `https://KULLANICI-ADIN.github.io/aydem-site/` adresinde durur.

---

## Önce bunları bil

- **Site herkese açık olur.** GitHub Pages ücretsiz planda depoyu da herkese açık ister. Depoda ve sitede görünen her şeyi (günlük üretim, ciro hesapları, `config.json`, `data/raw_daily.csv`) linki bilen herkes görebilir. Bu rakamlar şirket içi veya gizli ise yayınlamadan önce karar ver.
- **EPİAŞ şifreni asla dosyaya yazma.** Aşağıda "Secrets" bölümüne yazacaksın, orada gizli kalır.
- **Kur değerleri örnek.** `config.json` içindeki USD/TRY kurlarını kendi değerlerinle değiştirmeden site sarı bir "örnek kur" uyarısı gösterir. YEKDEM'li üç santralin cirosu kura bağlıdır.
- **İlk çalışma uzun sürer** (1 Ekim 2025'ten bugüne kadar tüm veriyi çeker, genelde 15–60 dakika). Sonraki günler birkaç dakikadır.

---

## Kurulum (bir kere yapılır)

**1. GitHub hesabı aç:** github.com adresinden ücretsiz kayıt ol.

**2. Yeni depo oluştur:** Sağ üstte **+** → **New repository**.
- Repository name: `aydem-site`
- **Public** seç
- "Add a README file" gibi kutuları **işaretleme**
- **Create repository**

**3. Dosyaları yükle:** Bu zip'i bilgisayarında aç. Yeni depo sayfasında **uploading an existing file** bağlantısına tıkla. Klasörün **içindeki** her şeyi (`docs`, `data`, `.github` klasörleri ve tüm dosyalar) tarayıcıya sürükle-bırak, sonra **Commit changes**.
- Windows'ta `.github` klasörü görünmüyorsa: Dosya Gezgini → Görünüm → Göster → **Gizli öğeler**.
- Yüklemeden sonra depoda `.github/workflows/update.yml` dosyasının olduğunu kontrol et. Yoksa: **Add file → Create new file**, isim kutusuna `.github/workflows/update.yml` yaz (eğik çizgiler klasör oluşturur), zip'teki dosyanın içeriğini yapıştır, commit et.

**4. EPİAŞ giriş bilgilerini gizli sakla:** Depoda **Settings → Secrets and variables → Actions → New repository secret**. İki tane ekle:
- Name: `EPTR_USERNAME`, Secret: EPİAŞ Şeffaflık kayıt e-postan
- Name: `EPTR_PASSWORD`, Secret: şifren

**5. Kurları gir:** Depoda `config.json` dosyasını aç, sağ üstteki **kalem** simgesine tıkla.
- `kur_yillik` içine 2025 ve 2026 için ilgili dönemin ortalama USD/TRY kurunu yaz
- Kur ay ay çok değişiyorsa `kur_aylik` içine `"2025-11": 42.0` gibi ay ay yaz (ay değeri yıllığın önüne geçer)
- `"kur_onaylandi": false` satırını `true` yap
- **Commit changes**

**6. Siteyi aç:** **Settings → Pages**. "Build and deployment" altında Source: **Deploy from a branch**, Branch: **main**, klasör: **/docs** → **Save**.

**7. Veriyi ilk kez çek:** Üstteki **Actions** sekmesi → soldan **Veriyi güncelle** → **Run workflow** → yeşil **Run workflow**. (Actions kapalı diyorsa "I understand my workflows, go ahead and enable them" düğmesine bas.) Bitince site dolmuş olur.

**8. Siteye git:** `https://KULLANICI-ADIN.github.io/aydem-site/` (Settings → Pages sayfasında da yazar, ilk yayında birkaç dakika sürebilir).

---

## Sonrasında

| Ne yapmak istiyorsun | Nasıl |
|---|---|
| Hiçbir şey | Her sabah 07:30'da kendiliğinden güncellenir |
| Hemen yenilemek | Actions → Veriyi güncelle → Run workflow |
| Kur veya YEKDEM fiyatı değişti | `config.json`'u düzenle. Bir sonraki çalışmada hesap yenilenir, veri tekrar çekilmez |
| Bir santral YEKDEM'e girdi | `config.json` → `yekdem` listesine `"SANTRAL ADI": fiyat` ekle |
| Yeni santral ekledim | `santraller` listesine ekle, sonra Run workflow ekranında **"Tüm veriyi baştan çek"** kutusunu işaretle (eski günleri de çekmesi için) |

---

## Sorun çıkarsa

**Actions'ta kırmızı çarpı var:** Çalışmanın üstüne tıkla, hata satırını oku. En sık üç neden:
- Şifre/e-posta yanlış: Secrets'ı düzelt
- `config.json` bozulmuş (virgül veya tırnak eksik): son düzenlemeyi geri al
- EPİAŞ, GitHub'ın yurt dışı sunucularından gelen istekleri yanıtlamıyor (zaman aşımı hatası). Bu durumda veriyi kendi bilgisayarında üretip yükleyebilirsin, aşağıya bak.

**Site "Veri yüklenemedi" diyor:** İlk güncelleme henüz bitmemiş ya da hata vermiş. Actions sekmesine bak.

**Sitede sarı uyarı var:** "Örnek kur" ise `config.json`'da `kur_onaylandi`'yı `true` yap. "Eksik veri" ise EPİAŞ'tan bazı parçalar alınamamış, sonraki çalışmada kendiliğinden tekrar denenir.

### Veriyi kendi bilgisayarında üretip yüklemek (Actions çalışmazsa)

```
py -3.13 -m pip install -r requirements.txt
set EPTR_USERNAME=mailadresin@ornek.com
set EPTR_PASSWORD=sifren
py -3.13 update_data.py
```

Oluşan şu dosyaları GitHub'da ilgili klasörlere yükle (**Add file → Upload files**):
`docs/data.json`, `docs/aydem_gunluk_karsilastirma.xlsx`, `data/raw_daily.csv`

### Siteyi bilgisayarında denemek

`docs` klasöründe cmd aç, `py -3.13 -m http.server 8000` yaz, tarayıcıda `http://localhost:8000` adresine git. (`index.html`'e çift tıklarsan veri yüklenmez, tarayıcı dosya okumayı engeller.)

---

## Dosyalar

| Dosya | Ne işe yarar |
|---|---|
| `config.json` | Santral listesi, YEKDEM fiyatları, kurlar. **Düzenleyeceğin tek dosya** |
| `update_data.py` | Veriyi çeker ve hesaplar |
| `aydem_core.py` | EPİAŞ bağlantısı ve hesap fonksiyonları |
| `.github/workflows/update.yml` | "Her sabah 07:30'da çalıştır" zamanlayıcısı |
| `docs/index.html` | Site sayfası |
| `docs/data.json`, `docs/*.xlsx` | Güncelleme sonrası oluşur: sitenin okuduğu veri ve Excel |
| `data/raw_daily.csv` | Günlük ham veri (tekrar çekmemek için saklanır) |
