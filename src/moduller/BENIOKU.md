# Modül Sözleşmesi

Bu klasördeki her modül **ayrı geliştirilir** ve ana programın (`src/App.jsx`)
hiçbir dosyasını import etmez. Modülün ihtiyaç duyduğu her şey **prop olarak** gelir.

Amaç: modülü yazan kişi ana programı hiç açmadan çalışabilsin, yazdığı kod
mevcut ekranları ve verileri hiçbir şekilde bozamasın.

---

## 1. Dosya düzeni

```
src/moduller/
  BENIOKU.md              <- bu dosya
  planlama/
    Planlama.jsx          <- giriş noktası: "export default function Planlama(props)"
    (istediğin kadar yardımcı dosya: parcalar.jsx, hesap.js, ...)
```

Kurallar:

- Klasörün **tek bir giriş dosyası** olur ve `export default` ile bir React bileşeni verir.
- Klasör içindeki dosyalar birbirini import edebilir (`./hesap.js` gibi).
- Klasör dışına import **yasak**. `../../App.jsx` yazma. İhtiyacın varsa prop olarak isteyeceksin.

## 2. Dil ve teknoloji

| Konu | Kural |
|---|---|
| Dil | **JavaScript + JSX** (`.jsx`). TypeScript yok — proje TS kurulu değil. |
| React | 18, fonksiyon bileşeni + hook. Sınıf bileşeni kullanma. |
| Stil | **Satır içi stil** (`style={{...}}`) veya proje sınıfları (`card`, `input`, `field-label`, `btn-ghost`, `pill`). Tailwind YOK. Yeni global CSS dosyası ekleme. |
| İkon | `lucide-react` (projede kurulu). |
| Excel | `xlsx` (projede kurulu) — ama tercihen `ui.excelIndir` kullan. |
| Veritabanı | Firestore'a **doğrudan erişme**. Sadece `api` üzerinden (aşağıda). |
| Yeni paket | **Ekleme.** Bir kütüphaneye ihtiyacın olursa önce sor — `package.json` değişikliği ayrı bir işlem. |
| Tarayıcı deposu | `localStorage` / `sessionStorage` kullanma. |
| İsimlendirme | Türkçe (projenin geri kalanı Türkçe). |

Grafik/Gantt gibi işler için hazır kütüphane yerine `div` veya satır içi `<svg>`
ile çiz — proje bugüne kadar bunu böyle yaptı, ek paket getirmiyor.

## 3. Bileşenin aldığı proplar

```jsx
export default function Planlama({ kullanici, yetki, api, ui, veri }) { ... }
```

### `kullanici`
`{ email, ... }` — giriş yapmış kişi.

### `yetki`
`"duzenle" | "goruntule"` — bu ekran için kullanıcının yetkisi.
**`"goruntule"` ise kaydet/sil düğmelerini gizle.** (Sunucu tarafı da ayrıca engeller,
ama kullanıcıya çalışmayan düğme gösterme.)

### `api` — veri katmanı

Modül **sadece kendi koleksiyonlarına** erişebilir. Planlama modülünün öneki
`planlama_`. Başka bir ada erişmeye çalışırsan fonksiyon hata fırlatır ve
Firestore güvenlik kuralları da ayrıca reddeder.

| Fonksiyon | Ne yapar |
|---|---|
| `api.dinle(ad, geriCagir)` | Canlı dinler. `geriCagir(liste)` her değişimde çağrılır. Geriye "durdur" fonksiyonu döner — `useEffect` içinde `return durdur`. |
| `api.ekle(ad, veri)` | Yeni kayıt. `olusturma` ve `olusturanEposta` otomatik eklenir. Geriye id döner. |
| `api.ekleNumarali(ad, evrakNo, veri)` | Evrak numarasını belge kimliği yapar — aynı numara iki kez kaydedilemez (10 kişi aynı anda kaydetse bile). |
| `api.guncelle(ad, id, veri)` | Günceller. `guncellemeTarihi` ve `guncelleyen` otomatik eklenir. |
| `api.sil(ad, id)` | Siler. |
| `api.topluYaz(islemler)` | Toplu işlem: `[{ tur:"ekle"\|"guncelle"\|"sil", koleksiyon, id?, veri? }]` |
| `api.sonrakiNo(kayitlar, onek)` | Sıradaki evrak numarasını üretir (`PLN-0001` → `PLN-0002`). |

Örnek:

```jsx
useEffect(() => {
  const durdur = api.dinle("planlama_isler", (liste) => setKayitlar(liste));
  return durdur;
}, []);
```

Hata yakalama: yetkisiz yazma denemesinde hata nesnesinde `err.yetkiHatasi === true`
olur — bu durumda kullanıcıya ayrıca mesaj gösterme, program zaten gösteriyor.

### `ui` — hazır bileşenler ve stiller

| Alan | İçerik |
|---|---|
| `ui.stil` | `satir, etiket, giris, tabloBaslik, tabloHucre, hucreGiris, dugme, anaDugme, belgeKutu, belgeEtiket, renk` |
| `ui.stil.renk` | `zemin, kart, kenar, vurgu, uyari, hata, iyi, soluk, yazi` |
| `ui.Pencere` | Fiş/modal penceresi. `acik, kapat, baslik, ikon, genislik, butonlar, children` |
| `ui.SecimPenceresi` | Arama + filtreli seçim penceresi (stok, cari vb. seçtirmek için) |
| `ui.SecimAlani` | Tıklanınca seçim penceresi açan kutu |
| `ui.Stat` | Sayı kartı — `label`, `value`, `highlight` |
| `ui.UyariPenceresi` | Uyarı/onay penceresi |
| `ui.excelIndir(veri, dosyaAdi, sayfaAdi)` | Nesne dizisini Excel'e indirir |
| `ui.renkliExcelIndir({...})` | Satır renkli Excel (dolgu renkli) |
| `ui.sablonIndir(basliklar, ornekler, dosyaAdi, sayfaAdi)` | Boş şablon üretir |
| `ui.yazdir(secenekler)` | **AS9100 antetli A4 form** basar (doküman no, sayfa x/y, imza alanları dahil) |
| `ui.tarih` | `bugun()` → `"2026-08-20"`, `tr(iso)` → `"20.08.2026"` |
| `ui.sayi` | `cevir(v)` → sayı, `tr(n)` → `"1.234,56"`, `tl(n)` → `"1.234,56 ₺"` |

`ui.yazdir` kullanımı — kolonlar `al(satir, index)` fonksiyonu ister, satırlar nesne dizisidir:

```jsx
ui.yazdir({
  belgeAdi: "ÜRETİM PLANI",
  dokumanKodu: "planlama",          // Form Ayarları'ndaki doküman no ile eşleşir (isteğe bağlı)
  ustBilgiler: [["Tarih", ui.tarih.tr(ui.tarih.bugun())], ["Kayıt", String(liste.length)]],
  kolonlar: [
    { baslik: "#", gen: "8mm", hiza: "ort", al: (r, i) => i + 1 },
    { baslik: "Parça", al: (r) => r.parcaAdi },
    { baslik: "Miktar", gen: "20mm", hiza: "sag", al: (r) => r.miktar },
  ],
  satirlar: liste,
  imzalar: ["Hazırlayan", "Onaylayan"],
});
```

`hiza` değerleri: `"ort"` (ortala), `"sag"` (sağa daya), boş (sola).

### `veri` — ana programdan gelen SALT OKUNUR veriler

| Alan | İçerik |
|---|---|
| `veri.makineler` | Makine kartları |
| `veri.takimlar` | Takım kartları |
| `veri.stokKartlari` | Stok kartları (kod, ad, birim, miktar) |
| `veri.cariler` | Cari/firma kartları |
| `veri.hammaddeler` | Hammadde takip kayıtları |
| `veri.talepler` | Satınalma talepleri |
| `veri.siparisler` | Satınalma siparişleri |
| `veri.projeler` | Proje kartları |
| `veri.kullanicilar` | Kullanıcı listesi |
| `veri.formAyarlari` | Antet/firma bilgileri, doküman numaraları |

Bunlar **okunur, yazılmaz.** Bu verilere yazman gerekiyorsa söyle — bağlantıyı
ana programda ben kurarım (örn. "planlamadan satınalma talebi açılsın").

## 4. Görsel kurallar

- Koyu tema: zemin `#142a30`, kart `#1b333c`, kenar `#2a4b52`, vurgu `#2dd4bf`.
- Genişlik: 1280px'de yatay kaydırma **olmamalı**. Geniş tablolar
  `<div style={{ overflowX: "auto" }}>` içine alınır.
- Izgara kullanırken `gridTemplateColumns: "repeat(2, minmax(0, 1fr))"` yaz —
  `"1fr 1fr"` yazma, uzun metinde dışarı taşıyor.
- Mobil: 820px altında program tabloları otomatik kart görünümüne çeviriyor;
  `<table>` kullanırsan bedava geliyor.

## 5. Teslim

İki yol var, ikisi de olur:

**A) Sadece dosyaları yolla (önerilen).**
`src/moduller/planlama/` klasörünü zip'leyip gönder. Ben paketleyip
terminale yapıştırılacak `.sh` dosyasına çeviririm, testten geçiririm.
Bu yolda hiçbir şey kurmana gerek yok.

**B) Paketi kendin üret.**
Depoyu klonla, dosyalarını `src/moduller/planlama/` içine koy, sonra:

```bash
bash modul-paketle.sh planlama
```

Bu betik değişiklikleri sıkıştırıp tek bir `.sh` dosyası üretir; o dosya
terminale yapıştırılıp çalıştırılır. (Betik depoda mevcut.)

## 6. Kontrol listesi

Teslimden önce:

- [ ] Klasör dışına import yok
- [ ] Yeni npm paketi yok
- [ ] Firestore'a doğrudan erişim yok (sadece `api`)
- [ ] Koleksiyon adları `planlama_` ile başlıyor
- [ ] `yetki === "goruntule"` iken yazma düğmeleri gizli
- [ ] 1280px'de yatay kaydırma yok
- [ ] `npm run build` hatasız (ya da bende deneyeceksem söyle)
