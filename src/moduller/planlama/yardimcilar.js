// =====================================================================
// Planlama modülü — ortak yardımcı fonksiyonlar ve sabitler
// Bu dosya sadece hesaplama/veri üretimi içerir, ekran (JSX) yoktur.
// =====================================================================

// Operasyon rotasında kullanılabilecek istasyon tipleri.
// "makineGerekli: true" olanlar için sistem, veri.makineler içinden
// tightestTolerance'ı karşılayan uygun makineleri önerir (hard-block).
export const OPERASYON_TIPLERI = [
  { kod: "GKK", ad: "GKK – Giriş Kalite Kontrol", makineGerekli: false, birim: "GKK BÖLÜMÜ" },
  { kod: "CNC3", ad: "CNC 3 Eksen İşleme", makineGerekli: true, sinif: "CNC3EKSEN", birim: "ÜRETİM SAHASI" },
  { kod: "CNCT", ad: "CNC Torna", makineGerekli: true, sinif: "CNCTORNA", birim: "ÜRETİM SAHASI" },
  { kod: "TASLAMA", ad: "Satıh Taşlama", makineGerekli: true, sinif: "TASLAMA", birim: "ÜRETİM SAHASI" },
  { kod: "ARAKONTROL", ad: "Ara Kontrol", makineGerekli: false, birim: "ÜRETİM SAHASI" },
  { kod: "YUZEYSEVK", ad: "Yüzey İşlem Sevk (Fason)", makineGerekli: false, fasonGerekli: true, birim: "TEDARİKÇİ SAHASI" },
  { kod: "FINAL", ad: "Final Kontrol", makineGerekli: false, birim: "FİNAL KONTROL" },
  { kod: "SEVK", ad: "Sevkiyat", makineGerekli: false, birim: "SEVKİYAT ALANI" },
];

export function operasyonBul(kod) {
  return OPERASYON_TIPLERI.find((o) => o.kod === kod) || null;
}

// "±0.01" / "0,01 mm" / "0.005" gibi metinlerden mikron cinsinden sayı üretir.
// VARSAYIM: değerler mm cinsinden giriliyor. Projenizde farklıysa (örn. inch)
// bu fonksiyonu güncelleyin — karşılaştırma mantığının geri kalanı etkilenmez.
export function toleransMikron(deger) {
  if (deger === undefined || deger === null || deger === "") return null;
  const temiz = String(deger).replace(",", ".").replace(/[^0-9.]/g, "");
  const sayi = parseFloat(temiz);
  if (Number.isNaN(sayi)) return null;
  return sayi * 1000;
}

// veri.makineler alan adları projeden projeye değişebilir; BENIOKU.md sadece
// "Makine kartları" der, kesin şema vermez. Olası isimleri sırayla dener.
// NOT: Gerçek alan adlarını paylaşırsanız burayı netleştiririm.
function makineAlan(m, adaylar, varsayilan) {
  for (const ad of adaylar) {
    if (m && m[ad] !== undefined && m[ad] !== null && m[ad] !== "") return m[ad];
  }
  return varsayilan;
}

// Bir operasyon için, verilen dar tolerans şartını karşılayan aktif makineleri
// veri.makineler içinden filtreler. AS9100 GD&T kilidi burada uygulanır:
// tolerans şartını karşılamayan / sınıf uyuşmayan makineler listeye HİÇ girmez.
export function makineOner(operasyonKod, tightestTolerance, makineler) {
  const op = operasyonBul(operasyonKod);
  if (!op || !op.makineGerekli) return { uygunlar: [], geregi: false };
  const gereken = toleransMikron(tightestTolerance);
  const liste = Array.isArray(makineler) ? makineler : [];
  const uygunlar = liste.filter((m) => {
    const sinif = makineAlan(m, ["sinif", "makineSinifi", "tip", "kategori"], "");
    const sinifTemiz = String(sinif).toUpperCase().replace(/[\s_]/g, "");
    if (!sinifTemiz.includes(op.sinif)) return false;

    const aktif = makineAlan(m, ["aktif", "durum"], true);
    if (aktif === false || aktif === "pasif" || aktif === "Pasif") return false;

    if (gereken === null) return true; // tolerans girilmemişse sınıf eşleşmesi yeterli
    const kapasite = toleransMikron(
      makineAlan(m, ["tightestTolerance", "hassasiyet", "minTolerans", "kapasiteTolerans"], null)
    );
    if (kapasite === null) return true; // makinede kapasite bilgisi yoksa engelleme, sadece uyarısız geçer
    return kapasite <= gereken; // makine, istenenden daha sıkı/eşit toleransa çıkabilmeli
  });
  return { uygunlar, geregi: true };
}

export function makineAdi(id, makineler) {
  if (!id) return "-";
  const m = (Array.isArray(makineler) ? makineler : []).find((mk) => mk.id === id);
  if (!m) return id;
  return m.ad || m.makineAdi || id;
}

// Sipariş kaydından "SiparişNo" üretir: ProjeKodu/KalemNo (örn. "2026-167/5").
export function siparisNoUret(siparis) {
  const proje = (siparis?.projeKodu || siparis?.parcaNo || "").trim();
  const kalem = (siparis?.kalemNo || "").trim();
  return kalem ? `${proje}/${kalem}` : proje;
}

// WBS kodu üretimi: kök -> "SiparişNo/1", altına eklenen -> "SiparişNo/1-1",
// onun altına -> "SiparişNo/1-1-1" ...
export function kokWbsKoduUret(siparis) {
  return `${siparisNoUret(siparis)}/1`;
}
export function altWbsKoduUret(parentKodu, kardesSayisi) {
  const sira = (kardesSayisi || 0) + 1;
  return `${parentKodu}-${sira}`;
}

// Rota kütüphanesi anahtarı: aynı parça+revizyon ikinci kez rota istediğinde
// aynı anahtara düşer -> api.ekleNumarali bunu tekilleştirir, sistem otomatik
// mevcut rotayı bulup sadece ONAY ister (adımlar yeniden girilmez).
export function rotaAnahtari(parcaNo, revizyon) {
  const p = (parcaNo || "").trim().toUpperCase().replace(/\s+/g, "");
  const r = (revizyon || "00").trim().toUpperCase().replace(/\s+/g, "");
  return `ROTA-${p}-REV${r}`;
}

const RENK_YESIL = "yesil";
const RENK_SARI = "sari";
const RENK_KIRMIZI = "kirmizi";
const RENK_GRI = "gri"; // n/a (bu kalem için ilgili değil)

// Bir WBS kaleminin Üretime Hazırlık Skoru'nu (M/S/F/Q) hesaplar.
// M: Malzeme  S: Alt Parça  F: Fason  Q: Kalite/FAI
export function hazirlikHesapla(node, tumWbs, malzemeTalepleri, fasonTalepleri) {
  // Malzeme (M)
  let malzeme = RENK_GRI;
  if (node.malzemeGerekli && node.malzemeKaynagi === "musteri") {
    // Hammadde müşteriden serbest (free-issue) geliyor — satınalma beklemeye gerek yok.
    malzeme = RENK_YESIL;
  } else {
    const ilgiliMalzeme = (malzemeTalepleri || []).filter((t) => t.wbsId === node.id);
    if (ilgiliMalzeme.length > 0) {
      if (ilgiliMalzeme.some((t) => t.durum === "karsilandi")) malzeme = RENK_YESIL;
      else if (ilgiliMalzeme.some((t) => t.durum === "siparişte" || t.durum === "gonderildi")) malzeme = RENK_SARI;
      else malzeme = RENK_KIRMIZI;
    } else if (node.malzemeGerekli) {
      malzeme = RENK_KIRMIZI;
    }
  }

  // Alt Parça (S)
  const cocuklar = (tumWbs || []).filter((w) => w.parentId === node.id);
  let altParca = RENK_GRI;
  if (cocuklar.length > 0) {
    if (cocuklar.every((c) => c.durum === "Tamamlandı")) altParca = RENK_YESIL;
    else if (cocuklar.some((c) => c.durum === "Üretimde" || c.durum === "Rotalandı")) altParca = RENK_SARI;
    else altParca = RENK_KIRMIZI;
  }

  // Fason (F)
  let fason = RENK_GRI;
  if (node.fasonGerekli) {
    const ilgiliFason = (fasonTalepleri || []).filter((t) => t.wbsId === node.id);
    if (ilgiliFason.some((t) => t.durum === "dondu")) fason = RENK_YESIL;
    else if (ilgiliFason.some((t) => t.durum === "sevk_edildi")) fason = RENK_SARI;
    else fason = RENK_KIRMIZI;
  }

  // Kalite / FAI (Q)
  let kalite = RENK_GRI;
  if (node.faiGerekli) {
    if (node.faiOnay === true) kalite = RENK_YESIL;
    else if (node.faiDurum === "beklemede") kalite = RENK_SARI;
    else kalite = RENK_KIRMIZI;
  }

  return { malzeme, altParca, fason, kalite };
}

export function planlanabilirMi(hazirlik) {
  return [hazirlik.malzeme, hazirlik.altParca, hazirlik.fason, hazirlik.kalite].every(
    (r) => r === "yesil" || r === "gri"
  );
}

// Sipariş içe aktarma şablonu (ui.sablonIndir ile üretilir)
export const SIPARIS_SABLON_BASLIKLAR = [
  "Musteri Adi", "Proje Kodu", "Kalem No", "Parca No", "Parca Adi",
  "Revizyon", "Miktar", "Teslim Tarihi", "SFC", "Dar Tolerans", "Gerekli Makine Sinifi",
];

export const SIPARIS_SABLON_ORNEK = [
  ["ERKUR", "2026-167", "5", "2795015008", "YUKSEKLIK AYAR RULMAN PLAKASI", "2026-167-5", 4, "2026-09-15", "HAYIR", "0.01", "CNC3"],
];

// =====================================================================
// PROJELENDİRME TAMAMLANMA KAPISI
// Sipariş, ancak WBS ağacındaki HER kalem için üç şart sağlandığında
// "Planlamaya Hazır" olur:
//   1) bomTamamlandi   -> kişi kırılımı bitirdiğini işaretlemiş (tek tık)
//   2) rota güncel      -> node.rotaId, o parça+revizyonun anahtarına eşit
//   3) malzeme kararı   -> gerekmiyor / müşteri malzemesi / talep açılmış
// Bu modül Çizelgeleme (gelecek modül) için tek doğruluk kaynağıdır;
// Çizelgeleme sadece bu fonksiyonun sonucunu / planlama_paketler'i okur.
// =====================================================================
export function projeTamamlanmaDurumu(siparisId, tumWbs, malzemeTalepleri) {
  const dugumler = (tumWbs || []).filter((w) => w.siparisId === siparisId);
  const toplam = dugumler.length;
  let kirilimTamam = 0, rotaTamam = 0, malzemeTamam = 0;

  dugumler.forEach((n) => {
    if (n.bomTamamlandi) kirilimTamam++;

    const anahtar = rotaAnahtari(n.parcaNo, n.revizyon);
    if (n.rotaId && n.rotaId === anahtar) rotaTamam++;

    const malzemeKararVerildi =
      !n.malzemeGerekli ||
      n.malzemeKaynagi === "musteri" ||
      (malzemeTalepleri || []).some((t) => t.wbsId === n.id);
    if (malzemeKararVerildi) malzemeTamam++;
  });

  const hepsiTamam = toplam > 0 && kirilimTamam === toplam && rotaTamam === toplam && malzemeTamam === toplam;
  return { toplam, kirilimTamam, rotaTamam, malzemeTamam, hepsiTamam };
}

// İki ISO zaman damgası arasındaki süreyi "X gün Y sa." / "X sa. Y dk." biçiminde yazar.
export function sureFormatla(baslangicIso, bitisIso) {
  if (!baslangicIso || !bitisIso) return null;
  const ms = new Date(bitisIso).getTime() - new Date(baslangicIso).getTime();
  if (Number.isNaN(ms) || ms < 0) return null;
  const toplamDakika = Math.floor(ms / 60000);
  const gun = Math.floor(toplamDakika / (60 * 24));
  const saat = Math.floor((toplamDakika % (60 * 24)) / 60);
  const dakika = toplamDakika % 60;
  if (gun > 0) return `${gun} gün ${saat} sa.`;
  return `${saat} sa. ${dakika} dk.`;
}

export function saatSayisi(baslangicIso, bitisIso) {
  if (!baslangicIso || !bitisIso) return null;
  const ms = new Date(bitisIso).getTime() - new Date(baslangicIso).getTime();
  if (Number.isNaN(ms) || ms < 0) return null;
  return Math.round((ms / 3600000) * 10) / 10;
}

// Ortak hata gösterimi: yetkisiz yazma denemesinde program zaten mesaj
// gösterdiği için (err.yetkiHatasi === true) burada sessiz geçilir.
export async function guvenliCagir(fn, hataBasligi) {
  try {
    await fn();
    return true;
  } catch (err) {
    if (err && err.yetkiHatasi) return false;
    alert(`${hataBasligi || "İşlem başarısız"}: ${err && err.message ? err.message : String(err)}`);
    return false;
  }
}
