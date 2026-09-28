# TradingView Hisse Tarama → Excel

TradingView'in **"Sütun ekle"** listesindeki tüm metrikleri otomatik keşfeden, seçilen
piyasadaki hisselerin verilerini çekip **tek bir Excel dosyasına** yazan tek dosyalık
bir Python aracı.

BIST örneğinde **650 sembol × 994 sütun** veriyi ~20 saniyede tabloya dönüştürüyor.

```
tradingview_screener.py      # tek dosya, tek bağımlılık (openpyxl)
```

---

## Öne çıkanlar

| | |
|---|---|
| **Sütun keşfi** | TradingView'in sütun listesi bir API'de değil, screener sayfasının JavaScript paketlerinin içinde. Araç bu paketleri tarayıp tanımları çıkarıyor. |
| **Otomatik genişletme** | Mali dönem, gösterge uzunluğu, zaman dilimi, büyüme tipi ve plot varyantları parametrelerden türetiliyor → 2.463 aday sütun. |
| **Akıllı eleme** | API'ye sorulan 2.463 sütunun 1.468'i BIST'te veri döndürmüyor ve atılıyor. Çıktıda boş sütun kalmıyor. |
| **Türkçe etiket** | 236 anahtar için elle yazılmış etiket + sözlük tabanlı otomatik çeviri (`total_revenue_fq` → `Toplam Gelir (Çeyreklik)`). |
| **Doğru sayı biçimi** | Yüzde, oran, para ve fiyat sütunları ayrı ayrı biçimlendirilir. |
| **Sektör çevirisi** | 123 sektör / alt sektör / ülke adı Türkçeleştirilir. |
| **Çift çıktı** | Her çalıştırma hem Excel (tarihli + sabit isimli kopya) hem de 21 MB'lık satır tabanlı JSON üretir. |
| **Önbellek** | Sütun kataloğu bir kez keşfedilip saklanıyor; sonraki çalıştırmalar ~20 saniye. |

## Gereksinimler

- Python 3.9 veya üzeri
- `openpyxl`

```bash
pip install openpyxl
```

Başka hiçbir şey gerekmiyor — `requests`, `pandas`, `bs4` gibi paketler kullanılmıyor.

## Hızlı başlangıç

```bash
python tradingview_screener.py
```

Her çalıştırma `Archive/` klasörüne üç dosya yazar:

| Dosya | Boyut | Açıklama |
|---|---|---|
| `BIST_20260928.xlsx` | 5,4 MB | Tarihli Excel — her çalıştırma ayrı dosya |
| `Bist_Hisse_Tum_Veriler.xlsx` | 5,4 MB | Sabit isimli kopya, her seferinde yenilenir |
| `Bist_Hisse_Tum_Veriler.json` | 21,6 MB | Aynı verinin JSON hâli |

Sabit isimli dosya her zaman son çalıştırmanın verisini tutar, böylece
`Bist_Hisse_Tum_Veriler.xlsx` yolunu bir PowerShell betiğine sabit yazabilirsiniz.

İlk çalıştırma sütun kataloğunu keşfettiği için ~2 dakika sürer. Sonraki
çalıştırmalar önbellekten yararlanıp ~20 saniyede biter.

## Komut satırı seçenekleri

| Seçenek | Açıklama |
|---|---|
| `--market KOD` | Piyasa kodu. Varsayılan `turkey` (BIST). Örnekler: `usa`, `germany`, `japan`, `india`, `hongkong`, `global` |
| `--cikti DOSYA` | Çıktı dosyasının tam yolu. Verilmezse `./Archive` klasörüne yazılır |
| `--sadece-hisse` | Sadece hisseleri alır, fon/ETF'leri çıkarır (650 → 626) |
| `--sutun A,B,C` | Virgülle ayrılmış sütun anahtarları. Boş bırakılırsa tümü |
| `--katalog-yenile` | Önbelleği yok sayıp sütunları yeniden keşfeder |
| `--katalog-yaz DOSYA` | Sütun kataloğunu incelenebilir bir CSV olarak da yazar |

```bash
# hızlı sütun listesi için
python tradingview_screener.py --katalog-yaz sutun_katalogu.csv

# sadece birkaç sütun
python tradingview_screener.py --sutun "name,close,market_cap_basic,Perf.Y,RSI"

# ABD piyasası
python tradingview_screener.py --market usa --cikti "ABD.xlsx"
```

## Excel çıktısı

**1. sayfa — BIST verisi**

- Üst satır: Türkçe etiket · Alt satır: ham TradingView anahtarı
- İlk 2 sütun (`BIST Kodu`, `Hisse Adı`) ve ilk 2 satır dondurulmuş
- Tüm sütunlarda AutoFilter açık
- Sayı biçimleri sütun tipine göre: yüzde `0.00"%"`, para `#,##0`, fiyat `0.000`, oran/teknik `0.00`

**2. sayfa — Sütun Kataloğu**

Her sütun için anahtar, Türkçe etiket, kategori ve kaç satırda veri olduğu. Hangi
metriğin neden geldiğini anlamak için ilk bakılacak yer burası.

Sütunlar 12 kategoriye ayrılır: Teknik, Teknik · Tavsiye, Fiyat · Hacim, Finansal,
Finansal · Büyüme, Değerleme, Temettü, Performans, Analist, Şirket, Kimlik, Fon / ETF.

> **Yüzdeler hakkında önemli:** TradingView yüzde sütunlarını `12.62` şeklinde
> döndürür, `0.1262` değil. Bu yüzden Excel'in yüzde biçimi 100 ile çarpılmayacak
> şekilde `0.00"%"` olarak ayarlanır — böylece hücrede `12.62%` görünür, değer
> Excel'in gizli yüzde mantığına yanlış girmez.

## JSON çıktısı

`Bist_Hisse_Tum_Veriler.json` aynı verinin makine tarafından okunabilir hâlidir.
Satır tabanlıdır — dizideki her nesne bir sembolü, içindeki her alan bir sütunu
temsil eder:

```json
[
  {
    "symbol": "THYAO",
    "ticker": "BIST:THYAO",
    "name": "THYAO",
    "sector": "Ulaşım",
    "industry": "Havayolları",
    "close": 287,
    "market_cap_basic": 401235000000,
    "dividends_yield": 0.8
  }
]
```

- `symbol` Excel'deki **BIST Kodu** sütunuyla birebir aynıdır
- `ticker` piyasa ön ekli tam TradingView sembolüdür
- Alan adları Excel'in alt satırındaki ham anahtarlarla aynıdır
- Eksik veriler `null` olarak yazılır, atlanmaz

pandas'ta doğrudan okunur:

```python
import pandas as pd
df = pd.read_json("Archive/Bist_Hisse_Tum_Veriler.json")
```

## Sütun keşfi nasıl çalışıyor?

TradingView'in tarayıcı uç noktası (`scanner.tradingview.com/{piyasa}/scan`) sadece
istediğiniz sütun adlarını verdiğinizde veri döndürür — hangi sütunların var
olduğunu söyleyen bir uç nokta yok. Sütun listesi, screener sayfasının yüklediği
JavaScript paketlerinde tanımlıdır. Araç:

1. Screener sayfasını indirir, yüklenen paketleri bulur (~47 adet).
2. Tanım sınıfı içeren paketleri tanır, her birinden `_key` ve `sortKey` değerlerini
   çıkarır (**468 temel anahtar**).
3. Her anahtarın hangi parametreyi taşıdığını (`_getColumnKey` gövdesi) okuyup
   varyantları üretir.
4. Adayları toplu isteklerle (150 sütun/istek) API'ye sorar.
5. Hiç veri döndürmeyenleri eler.

### Doğrulanmış adlandırma kuralları

TradingView'in isimlendirmesi tutarsız olduğu için varyantlar tahmin edilmedi,
**API'ye tek tek sorularak teyit edildi**. Yanlış olanlar kodda bilerek yok:

| Tür | Geçerli | Geçersiz (kullanılmıyor) |
|---|---|---|
| Mali dönem | `total_revenue_fy`, `_fq`, `_ttm` | `total_revenue_annual` |
| Gösterge uzunluğu | `ADX_9`, `Mom_14` (alt çizgi) | `ADX\|14` |
| Hareketli ortalama | `SMA20`, `EMA20` (bitişik) | `SMA_20` |
| Zaman dilimi | `RSI\|1`, `\|5`, `\|15`, `\|30`, `\|60`, `\|120`, `\|240`, `\|1W`, `\|1M` | `RSI\|1D`, `\|3M`, `\|6M`, `\|1Y` |
| Performans | `Perf.W`, `.1M`, `.3M`, `.6M`, `.Y`, `.3Y`, `.5Y`, `.10Y`, `.YTD`, `.All` | `Perf.1Y`, `Perf.2Y` |
| Piyasa değeri performansı | `Perf.1W`, `.1M`, `.1Y`, `.5Y` | — (farklı adlandırma) |
| Değişim | `change\|1`, `\|5`, `\|15`, `\|30`, `\|60`, `\|120`, `\|240`, `\|1W`, `\|1M` | `change\|3M`, `\|6M`, `\|1Y`, `\|YTD`, `\|All` |

## Etiketleri özelleştirme

İki sözlük dosyanın içindedir; düzenleyip betiği tekrar çalıştırmak yeterli.

```python
LABEL_MAP = {
    "close": "Son Fiyat",
    "dividends_yield": "Temettü Verimi (%)",
    # ... kendi etiketini ekle
}

SECTOR_MAP = {
    "Turkey": "Türkiye",
    "Transportation": "Ulaşım",
    # ...
}
```

Listelenmeyen bir anahtar varsa `humanize()` sözlük tabanlı olarak otomatik
Türkçeleştirir (`net_debt_to_equity_fq` → `Net Borç / Özkaynak (Çeyreklik)`).
Aynı etiket iki sütunda çakışırsa ayırt edici bir ek otomatik eklenir, böylece
başlıklar hiçbir zaman tekrar etmez.

## Otomasyon: her gün otomatik veri çekme

Veriyi çeken kod Python, ama tetikleyici **Google Apps Script**. İki parça birbirine
bağlıdır:

```
Google Apps Script  ──►  GitHub Actions  ──►  python tradingview_screener.py
  (saat 19:00)             (bulutta çalışır)     (veriyi çeker, Excel'i depoya yazar)
```

Google Script veriyi kendisi çekmez; GitHub'a "veri-cek" olayı gönderir, Actions
çalıştırır. **Saat seçimi Script panelinde yapılır**, kodda sabit zaman yoktur.

### 1) Depoyu GitHub'a yükleyin

`github.com/new` → depoyu **Public** olarak oluşturun (README/.gitignore eklemeyin,
dosyalar zaten var). Sonra:

```bash
git init
git add .
git commit -m "ilk sürüm"
git branch -M main
git remote add origin https://github.com/KULLANICI/DEPO.git
git push -u origin main
```

`.github/workflows/fetch-data.yml` push edildiği anda Actions tarafından tanınır.

### 2) GitHub jetonu oluşturun

**Settings → Developer settings → Personal access tokens → Fine-grained tokens → New**

- **Repository access:** yalnız bu depo
- **Repository permissions → Contents:** `Read and write`

### 3) Google Apps Script kurun

`script.google.com` → **Yeni proje**:

1. `google-apps-script/Code.gs` içeriğini yapıştırın
2. `AYAR` bloğundaki `KULLANICI` ve `DEPO` alanlarını doldurun
3. **Project Settings → Script Properties** → anahtar `GH_PAT`, değer jetonunuz
4. `google-apps-script/appsscript.json` içeriğini **Show appsscript.json manifest file**
   ile açıp yapıştırın
5. `BILGIYI_KONTROL_ET()` fonksiyonunu bir kez çalıştırın; **Executions** çıktısında
   `Depo yanıtı: HTTP 200` görmelisiniz

### 4) Tetikleyici ekleyin

Sol menü → **Triggers** → **Add Trigger**

| Alan | Değer |
|---|---|
| Function | `gunlukVeri` |
| Deployment | `Head` |
| Trigger source | Time-driven |
| Type | Day timer |
| Time | `19:00` |

Birden fazla saat isterseniz adımı her saat için tekrarlayın; her biri ayrı tetikleyici
olur. İlk çalıştırmayı beklemeden Actions sekmesinden **Run workflow** ile elle
tetikleyip doğrulayabilirsiniz.

### Depoda ne birikir

Her çalıştırma `Archive/` altında üç dosya üretir, ancak **depoya yalnızca tarihli Excel
girer**:

| Dosya | Depoda | Neden |
|---|---|---|
| `BIST_20260928.xlsx` | saklanır | Günün verisi, tarih geçmişi |
| `Bist_Hisse_Tum_Veriler.xlsx` | `.gitignore` | Tarihli dosyanın birebir kopyası |
| `Bist_Hisse_Tum_Veriler.json` | `.gitignore` | Her gün 21,6 MB, tarihsel değeri yok |

Sabit isimli dosyalar her çalıştırmada yeniden üretilir; yalnız geçmişe gömülmezler.
Böylece depoya günlük ~5,4 MB yazılır (ayda ~160 MB) — normal git geçmişine
koyulsalardı yılda 6 GB'ı bulurdu.

Aynı gün birden fazla tetikleme gelirse numara eklenir (`BIST_20260928_2.xlsx`);
`concurrency` ayarı yüzünden çalıştırmalar birbirine girmez.

### Depo boyutunu kontrol etmek

```bash
git count-objects -vH          # depo ne kadar büyümüş
```

Eski tarihli dosyaları silip geçmişten de temizlemek mümkündür; pratikte yılda birkaç
kere `Archive/` altındaki eski dosyaları toplu silip normal `git commit` + `git push`
yapmak yeterlidir.

### Bilmeniz gerekenler

- **GitHub programlı işleri erteleyebilir.** 19:00 yerine 19:10–19:20'de çalışabilir.
- **60 gün hareketsiz kalan repolarda** zamanlanmış işler otomatik kapanır ve uyarı
  e-postası gelir. Bu iş her gün commit attığı için muhtemelen sayılır, yine de ara
  sıra Actions sekmesine bakın.
- **Jetonu `Code.gs` içine yazmayın.** Dosya repoya gidiyor; jeton Script Properties'te
  kalsın.
- Bu akış TradingView'ın resmî olmayan uç noktasını kullanır. Erişim kısıtlarsa iş
  kırmızıya döner ve Google size e-posta gönderir.

## Depo yapısı

```
tradingview_screener.py                       veriyi çeken tek dosya
requirements.txt                              bağımlılıklar (openpyxl)
.github/workflows/fetch-data.yml              GitHub Actions işi
google-apps-script/Code.gs                    tetikleyici (Apps Script)
google-apps-script/appsscript.json            manifest (saat dilimi: Europe/Istanbul)
.gitignore / .gitattributes                    depoya ne girer / satır sonları
README.md                                     bu dosya
Archive/                                      çıktılar (yalnız tarihli Excel commit edilir)
  BIST_20260928.xlsx                          tarihli Excel
  Bist_Hisse_Tum_Veriler.xlsx                 sabit isimli kopya
  Bist_Hisse_Tum_Veriler.json                 sabit isimli JSON
~/.tradingview_screener/sutun_katalogu.json   sütun önbelleği
```

Tarihli Excel `PİYASA_YYYYMMDD.xlsx` olarak adlandırılır. Aynı gün tekrar
çalıştırırsanız numara eklenir (`BIST_20260928_2.xlsx`) — hiçbir çalıştırma veri
kaybetmez. Sabit isimli iki dosya ise her seferinde üzerine yazılarak daima son
çalıştırmanın verisini tutar.

Önbellek konumu `TRV_HOME` ortam değişkeniyle değiştirilebilir.

## Sınırlar ve notlar

- Bu araç TradingView'in **resmî olmayan**, herkese açık tarayıcı uç noktasını
  kullanır. Login gerektirmez, istekler arası kısa beklemeler vardır.
- Sütun keşimi frontend paketlerine dayanır. TradingView paket yapısını değiştirirse
  araç net bir hata verir (sessizce bozulmaz) ve `--katalog-yenile` ile kurtarılabilir.
- Görülen tüm veriler TradingView'a aittir; kullanım şartlarına uymak sizindir.
- Bu araç verileri **görüntüler**, yatırım tavsiyesi üretmez.

## Lisans

MIT
