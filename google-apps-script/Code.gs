/**
 * Günlük BIST verisi tetikleyicisi
 * ================================
 *
 * Bu script kendisi veri çekmez. TradingView'dan veriyi çeken Python kodu
 * GitHub Actions'ta çalışır; bu script her gün o işi tetikler.
 *
 * Ne yapar
 * --------
 * GitHub'a "veri-cek" olayı gönderir. GitHub Actions bunu alır, Python'u
 * çalıştırır, sonucu depoya kaydeder. Saatin ne olduğu burada değil,
 * Google Script panelinde ayarladığınız tetikleyicide belirlenir.
 *
 * Kurulum
 * -------
 * 1) script.google.com -> Yeni proje
 * 2) Bu dosyanın içeriğini Code.gs yerine yapıştırın
 * 3) AYAR bölümünü kendi bilgilerinizle doldurun
 * 4) appsscript.json dosyasının içeriğini Dosya ayarları (Project Settings)
 *    -> "Göster: manifest dosyası" ile açıp yapıştırın
 * 5) Önce BILGIYI_KONTROL_ET() fonksiyonunu bir kez çalıştırın (token'ı doğrular)
 * 6) Sol menüden "Tetikleyiciler" (Triggers) -> "Tetikleyici ekle":
 *       - Hangi fonksiyon: gunlukVeri
 *       - Hangi kaynak: Saat bazlı (Time-driven)
 *       - Tetikleyici türü: Günde bir kez
 *       - Saat: istediğiniz saat (örn. 19:00)
 *    Birden fazla saat istiyorsanız bu adımı tekrarlayın; her saat için
 *    ayrı bir tetikleyici oluşur.
 *
 * Hata bildirimi
 * --------------
 * Tetikleyiciler hata olursa Google e-posta gönderir. Ek olarak hata
 * bildirimi webhook adresi tanımlarsanız (DISCORD_WEBHOOK) oraya da düşer.
 *
 * Not
 * ---
 * Bu script herkese açık bir GitHub API uç noktasına istek atar; kimlik
 * doğrulama için repository_dispatch yetkili bir kişisel erişim jetonu (PAT)
 * kullanır. Jetonu kaynak koda gömmeyin, script ayarlarından gizli değişken
 * olarak tanımlayın.
 */

// ======================================================================
// AYAR — kendi bilgilerinizle doldurun
// ======================================================================
var AYAR = {
  // GitHub kullanıcı adınız (projenin sahibi olan hesap)
  KULLANICI: "KULLANICI_ADINIZ",

  // Depo adı (uzantısız)
  DEPO: "DEPO_ADINIZ",

  // Depo varsayılan dalı
  DAL: "main",

  // GitHub kişisel erişim jetonu (PAT).
  // Oluşturma: GitHub -> Ayarlar -> Geliştiriciler ayarları ->
  //            Kişisel erişim jetonları -> Fine-grained token
  //            Depo: yalnız bu repo, İzinler -> Contents: Read and write
  //
  // Jetonu buraya düz yazmak yerine script ayarlarından gizli değişken
  // tanımlamanız önerilir. Gizli değişken adı: GH_PAT
  JETON: PropertiesService.getUserProperties().getProperty("GH_PAT")
};

// ======================================================================
// Ana tetikleyici
// ======================================================================

/**
 * Her gün (veya her saat) bu işi tetikler. Tetikleyici buraya bağlanır.
 */
function gunlukVeri() {
  tetikleVeriCek();
}

/**
 * GitHub'a "veri-cek" olayı gönderir. Tetikleyiciden bağımsız olarak
 * elle de çalıştırılabilir.
 */
function tetikleVeriCek() {
  var ayar = kontrolAyarlarini();
  var adres = "https://api.github.com/repos/" +
              ayar.KULLANICI + "/" + ayar.DEPO + "/dispatches";

  var yanit = UrlFetchApp.fetch(adres, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + ayar.JETON,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    },
    payload: JSON.stringify({ event_type: "veri-cek" }),
    muteHttpExceptions: true
  });

  var kod = yanit.getResponseCode();
  // 204 = kabul edildi. repository_dispatch gövdesiz yanıt döner.
  if (kod !== 204) {
    throw new Error(
      "GitHub isteği başarısız (HTTP " + kod + "): " + yanit.getContentText()
    );
  }
  return kod;
}

// ======================================================================
// Yardımcılar
// ======================================================================

/** Eksik ayarları yakalar; hata mesajı anlaşılır olsun diye. */
function kontrolAyarlari() {
  var ayar = AYAR;
  var eksik = [];

  if (!ayar.KULLANICI || ayar.KULLANICI === "KULLANICI_ADINIZ") {
    eksik.push("KULLANICI (GitHub kullanıcı adı)");
  }
  if (!ayar.DEPO || ayar.DEPO === "DEPO_ADINIZ") {
    eksik.push("DEPO (depo adı)");
  }
  if (!ayar.JETON) {
    eksik.push("GH_PAT (gizli değişken, Project Settings > Script properties)");
  }
  if (eksik.length > 0) {
    throw new Error("Eksik ayar: " + eksik.join(", "));
  }
  return ayar;
}

/**
 * Ayar ve jeton doğrulaması. Token'ı test etmek için önce bunu çalıştırın.
 * Sayfanın solundaki "Günlükte kayıt" (Executions) çıktısına bakın.
 */
function BILGIYI_KONTROL_ET() {
  var satirlar = [];
  try {
    var ayar = kontrolAyarlari();
    satirlar.push("Kullanıcı: " + ayar.KULLANICI);
    satirlar.push("Depo:    " + ayar.DEPO);
    satirlar.push("Jeton:   " + (ayar.JETON ? "tanımlı" : "YOK"));

    // Depoya erişim var mı, yalnız okumayla denetle
    var dogrula = UrlFetchApp.fetch(
      "https://api.github.com/repos/" + ayar.KULLANICI + "/" + ayar.DEPO,
      {
        headers: {
          Authorization: "Bearer " + ayar.JETON,
          Accept: "application/vnd.github+json"
        },
        muteHttpExceptions: true
      }
    );
    satirlar.push("Depo yanıtı: HTTP " + dogrula.getResponseCode());
  } catch (hata) {
    satirlar.push("HATA: " + hata.message);
  }
  console.log(satirlar.join("\n"));
}
