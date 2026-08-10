# Üretim Takip — Kurulum ve Yayınlama Rehberi

Bu proje: React (Vite) + Firebase Firestore (veritabanı) + Firebase Hosting (yayın).
Birkaç kişi ortak şifre ile girer, veriler herkeste anlık senkron güncellenir.

---

## 1. Firebase projesi oluştur

1. https://console.firebase.google.com adresine git, Google hesabınla giriş yap.
2. **Proje ekle** (Add project) → proje adı ver (örn. `uretim-takip`) → devam et → Google Analytics'i kapatabilirsin → **Proje oluştur**.

## 2. Firestore Database'i aç

1. Sol menüden **Build > Firestore Database**.
2. **Veritabanı oluştur** (Create database).
3. Konum seç (örn. `eur3` - Avrupa) → ileri.
4. Mod olarak **test modunda başlat** (test mode) seç. (Not: test modu 30 gün sonra kapanır, aşağıda kalıcı kural veriyorum.)

Firestore açıldıktan sonra **Rules** sekmesine gidip aşağıdakini yapıştır ve **Yayınla**:

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /{document=**} {
      allow read, write: if true;
    }
  }
}
```

> Not: Bu kural herkese açık okuma/yazma demektir — programın şifre kapısı sadece arayüz seviyesinde. Gerçek güvenlik istersen ileride Firebase Authentication'a geçeriz, şimdilik iç kullanım için yeterli.

## 3. Web uygulaması ekle ve config al

1. Proje ana sayfasında **</> (Web)** simgesine tıkla.
2. Uygulama adı gir (örn. `uretim-takip-web`) → **Uygulamayı kaydet**.
3. Karşına çıkan `firebaseConfig` nesnesindeki değerleri kopyala.
4. Bu projede `src/firebase.js` dosyasını aç, `BURAYA_...` yazan yerleri kendi değerlerinle değiştir.

## 4. Hosting'i etkinleştir

1. Sol menüden **Build > Hosting** → **Başlayın** (Get started) — sihirbazı adım adım geçebilirsin, CLI komutlarını aşağıda zaten vereceğim.

## 5. Ortak şifreyi ayarla

`src/config.js` dosyasını aç, `APP_PASSWORD` değerini ekibin kullanacağı şifreyle değiştir.

## 6. Bilgisayarında ilk çalıştırma (test için)

Node.js kurulu olmalı (https://nodejs.org). Sonra proje klasöründe:

```bash
npm install
npm run dev
```

Terminalde çıkan `http://localhost:xxxx` adresini tarayıcıda aç, şifreni gir, test et.

## 7. GitHub'a yükle

```bash
git init
git add .
git commit -m "İlk sürüm"
```

GitHub'da yeni boş bir repo oluştur (github.com → New repository), sonra:

```bash
git remote add origin https://github.com/KULLANICI_ADIN/REPO_ADI.git
git branch -M main
git push -u origin main
```

## 8. Firebase CLI kur ve giriş yap

```bash
npm install -g firebase-tools
firebase login
```

Tarayıcı açılacak, Firebase hesabınla onayla.

## 9. Projeyi Firebase'e bağla

Proje klasöründe:

```bash
firebase init hosting
```

Sorulara şu şekilde cevap ver:
- "Use an existing project" → oluşturduğun Firebase projesini seç
- "What do you want to use as your public directory?" → `dist` yaz
- "Configure as a single-page app?" → **Yes**
- "Set up automatic builds with GitHub?" → **Yes** dersen her `git push` sonrası otomatik yayınlar (önerilir); hesap bağlamanı ister, ekrandaki adımları takip et.

## 10. Yayınla

```bash
npm run build
firebase deploy
```

Terminalde sana bir **Hosting URL** verecek (örn. `https://uretim-takip.web.app`) — programın canlı adresi bu olacak.

GitHub Actions'ı kurduysan bundan sonra sadece `git push` yapman yeterli, otomatik build alıp yayınlayacak.

---

## Sıradaki geliştirmeler için hazır

- Makine listesini bana gönder, `Makineler` sekmesine toplu ekleyelim.
- İstersen raporlara grafik (günlük/haftalık üretim trendi) ekleyebiliriz.
- İstersen Excel/CSV dışa aktarma ekleyebiliriz.
- İstersen ileride her kullanıcıya ayrı giriş (Firebase Authentication) geçebiliriz.
