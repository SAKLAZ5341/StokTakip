// =====================================================================
// ÇİZELGELEME MOTORU (Planlama alt-modülü)
// Kullanıcının APS makrosundaki mantığın web/canlı-veri karşılığı:
//   - PerformStrategicSorting  -> oncelikSirala()
//   - dictQueues (tezgah kuyruğu) -> makineKuyrugu (bu dosyada, her hesapta sıfırdan)
//   - dictJobReady (bir önceki adımın bitişi) -> adimHazirTarihi
//   - MONTAJ 3 geçişli yakınsama -> hepsiniCizelgele() içindeki 3 pas
//   - ComputeEndDate_LTSL (vardiya takvimi) -> saatEkle() [BASİTLEŞTİRİLMİŞ:
//     tam vardiya/gece-mesaisi matrisi veri kaynağı (veri.vardiyalar) şu an
//     modül sözleşmesinde yok; bu yüzden planlama_ayarlar.gunlukKapasiteSaat
//     ile "iş günü = X saat" basit takvimi kullanılıyor. Gerçek vardiya
//     matrisi eklenmek istenirse bu tek fonksiyon değişir, geri kalanı sabit.]
//
// KRİTİK TASARIM KARARI: Hiçbir başlangıç/bitiş tarihi Firestore'a YAZILMAZ.
// Kullanıcının notu: "ham madde gelmedikçe iş termini dinamik ötelenir" —
// bu ancak tarihler her defasında YENİDEN hesaplanırsa doğru olur. Bu yüzden
// bu dosyadaki her şey SAF (yan etkisiz) fonksiyondur; UI her render'da
// (veya ilgili koleksiyon değiştiğinde) yeniden çağırır.
// =====================================================================

export const IS_EMRI_TURLERI = [
  { kod: "SERI", ad: "Seri", oncelik: 3 },
  { kod: "FAI", ad: "FAI (İlk Ürün Muayenesi)", oncelik: 1 },
  { kod: "DELTA", ad: "Delta (Kalan Miktar)", oncelik: 2 },
  { kod: "FAI_REWORK", ad: "FAI Rework", oncelik: 1 },
  { kod: "MONTAJ", ad: "Montaj", oncelik: 3 },
];

export function isEmriTuruBul(kod) {
  return IS_EMRI_TURLERI.find((t) => t.kod === kod) || IS_EMRI_TURLERI[0];
}

const VARSAYILAN_AYARLAR = { gunlukKapasiteSaat: 8 };

// İş saatini takvim tarihine ekler. v1: hafta sonu/vardiya matrisi yok,
// sadece "günlük kapasite saat"e göre orantılı takvim günü ekler.
export function saatEkle(baslangicIso, saat, ayarlar) {
  const gunlukSaat = (ayarlar && ayarlar.gunlukKapasiteSaat) || VARSAYILAN_AYARLAR.gunlukKapasiteSaat;
  const guvenliSaat = Number.isFinite(saat) && saat > 0 ? saat : 0.1;
  const gun = guvenliSaat / gunlukSaat;
  const ms = gun * 24 * 3600 * 1000;
  return new Date(new Date(baslangicIso).getTime() + ms).toISOString();
}

// Bir WBS kaleminin hammaddesi ne zaman "hazır" sayılır?
// - malzeme gerekmiyorsa: hemen (null tarih = "şimdi" anlamına gelir, çağıran yorumlar)
// - müşteri malzemesi: hemen hazır kabul edilir (free-issue)
// - satınalma: talep karşılandıysa fiiliTeslimTarihi; yoldaysa tahminiTeslimTarihi (belirsiz
//   işaretiyle); hiçbiri yoksa "belirsiz" -> iş emrinin termini hesaplanamaz, ötelenir.
export function hammaddeHazirTarihi(node, malzemeTalepleri) {
  if (!node || !node.malzemeGerekli) {
    return { tarih: null, belirsiz: false, aciklama: "Malzeme gerekmiyor" };
  }
  if (node.malzemeKaynagi === "musteri") {
    return { tarih: null, belirsiz: false, aciklama: "Müşteri malzemesi (free-issue) — hazır kabul edilir" };
  }
  const talepler = (malzemeTalepleri || []).filter((t) => t.wbsId === node.id);
  const karsilanan = talepler.find((t) => t.durum === "karsilandi" && t.fiiliTeslimTarihi);
  if (karsilanan) {
    return { tarih: karsilanan.fiiliTeslimTarihi, belirsiz: false, aciklama: "Hammadde teslim alındı" };
  }
  const yolda = talepler.find((t) => (t.durum === "siparişte" || t.durum === "gonderildi") && t.tahminiTeslimTarihi);
  if (yolda) {
    return { tarih: yolda.tahminiTeslimTarihi, belirsiz: true, aciklama: "Tahmini tedarikçi teslim tarihi (kesinleşmedi)" };
  }
  return { tarih: null, belirsiz: true, aciklama: "Hammadde henüz temin edilmedi — termin hesaplanamıyor" };
}

function oncelikPuani(isEmri, node) {
  const tur = isEmriTuruBul(isEmri.tur);
  let puan = tur.oncelik;
  if (node && node.faiGerekli && !node.faiOnay) puan -= 0.5; // FAI bekleyen işler öne alınır
  return puan;
}

function terminSirala(isEmirleri, wbsMapById, siparisMapById) {
  return [...isEmirleri]
    .filter((ie) => ie.durum !== "İptal" && ie.durum !== "Tamamlandı")
    .sort((a, b) => {
      const na = wbsMapById[a.wbsId], nb = wbsMapById[b.wbsId];
      const pa = oncelikPuani(a, na), pb = oncelikPuani(b, nb);
      if (pa !== pb) return pa - pb;
      const ta = (siparisMapById[a.siparisId] && siparisMapById[a.siparisId].teslimTarihi) || "9999-12-31";
      const tb = (siparisMapById[b.siparisId] && siparisMapById[b.siparisId].teslimTarihi) || "9999-12-31";
      if (ta !== tb) return ta < tb ? -1 : 1;
      return String(a.olusturma || a.id || "").localeCompare(String(b.olusturma || b.id || ""));
    });
}

// Bir tezgahın şu anki toplam kuyruk yükünü (saat) ve müsait olacağı tarihi
// döner — istasyon SEÇERKEN "kapasite doluluk" uyarısı için kullanılır.
// NOT: Bu son hesaplanmış çizelgeye bakar (hepsiniCizelgele sonucu), canlıdır.
export function makineDolulukOzeti(makineId, cizelgeSonucu) {
  let sonBitis = null;
  let toplamSaat = 0;
  let isSayisi = 0;
  Object.values(cizelgeSonucu.isEmriSonuclari || {}).forEach((sonuc) => {
    (sonuc.adimlar || []).forEach((a) => {
      if (!a.atlandi && a.makineId === makineId) {
        toplamSaat += a.sureSaat || 0;
        isSayisi += 1;
        if (!sonBitis || new Date(a.bitis) > new Date(sonBitis)) sonBitis = a.bitis;
      }
    });
  });
  return { sonBitis, toplamSaat, isSayisi };
}

// =====================================================================
// ANA MOTOR: tüm aktif iş emirlerini önceliğe göre sırayla çizelgeler.
// MONTAJ bağımlılığı için makro ile aynı yaklaşım: 3 geçiş, kuyruklar her
// geçişte sıfırdan kurulur, alt parça bitişleri bir önceki geçişten öğrenilir.
// =====================================================================
export function hepsiniCizelgele(tumIsEmirleri, tumWbs, malzemeTalepleri, tumSiparisler, ayarlar) {
  const wbsMapById = {};
  (tumWbs || []).forEach((w) => { wbsMapById[w.id] = w; });
  const siparisMapById = {};
  (tumSiparisler || []).forEach((s) => { siparisMapById[s.id] = s; });

  const sirali = terminSirala(tumIsEmirleri || [], wbsMapById, siparisMapById);

  let altParcaMinBaslangic = {}; // node.id (montajın kendi düğümü) -> en geç alt parça bitişi

  let sonIsEmriSonuclari = {};
  const GECIS_SAYISI = 3;

  for (let gecis = 0; gecis < GECIS_SAYISI; gecis++) {
    const makineKuyrugu = {}; // makineId -> son bitiş ISO
    const isEmriSonuclari = {};
    const yeniAltParcaMinBaslangic = {};

    sirali.forEach((ie) => {
      const node = wbsMapById[ie.wbsId];
      if (!node) return;

      const hazir = hammaddeHazirTarihi(node, malzemeTalepleri);
      let hazirTarih = hazir.tarih || new Date().toISOString();
      let belirsiz = hazir.belirsiz;

      // MONTAJ: alt parçaların (bir önceki geçişten bilinen) en geç bitişini de bekler.
      if (ie.tur === "MONTAJ" && altParcaMinBaslangic[node.id]) {
        if (altParcaMinBaslangic[node.id] > hazirTarih) hazirTarih = altParcaMinBaslangic[node.id];
      }

      let onceki = hazirTarih;
      const adimSonuclari = [];

      (ie.adimlar || []).forEach((adim) => {
        if (adim.dahil === false) {
          adimSonuclari.push({ ...adim, baslangic: null, bitis: null, atlandi: true });
          return;
        }
        const sure = adim.sureSaat || 0;
        let baslangic = onceki;

        if (adim.makineGerekli && adim.makineId) {
          const kuyrukBitis = makineKuyrugu[adim.makineId];
          if (kuyrukBitis && kuyrukBitis > baslangic) baslangic = kuyrukBitis;
        }

        const bitis = saatEkle(baslangic, sure, ayarlar);
        adimSonuclari.push({ ...adim, baslangic, bitis, atlandi: false });

        if (adim.makineGerekli && adim.makineId) {
          makineKuyrugu[adim.makineId] = bitis;
        }
        onceki = bitis;
      });

      const gecerliAdimlar = adimSonuclari.filter((a) => !a.atlandi);
      const genelBitis = gecerliAdimlar.length ? gecerliAdimlar[gecerliAdimlar.length - 1].bitis : hazirTarih;

      const siparis = siparisMapById[ie.siparisId];
      let terminDurumu = "Termin Yok";
      let gecikmeGun = 0;
      if (siparis && siparis.teslimTarihi) {
        const terminMs = new Date(siparis.teslimTarihi).getTime();
        const bitisMs = new Date(genelBitis).getTime();
        if (bitisMs <= terminMs) {
          terminDurumu = belirsiz ? "Belirsiz (Hammadde Bekleniyor)" : "Zamanında";
        } else {
          terminDurumu = "Gecikiyor";
          gecikmeGun = Math.ceil((bitisMs - terminMs) / (24 * 3600 * 1000));
        }
      } else if (belirsiz) {
        terminDurumu = "Belirsiz (Hammadde Bekleniyor)";
      }

      isEmriSonuclari[ie.id] = {
        isEmriId: ie.id, wbsId: ie.wbsId,
        adimlar: adimSonuclari, genelBaslangic: hazirTarih, genelBitis,
        belirsiz, terminDurumu, gecikmeGun,
        malzemeAciklama: hazir.aciklama,
      };

      // Bu iş bir MONTAJ'ın çocuğu ise (parentId), o montajın min başlangıcını güncelle.
      if (node.parentId) {
        const mevcut = yeniAltParcaMinBaslangic[node.parentId];
        if (!mevcut || genelBitis > mevcut) yeniAltParcaMinBaslangic[node.parentId] = genelBitis;
      }
    });

    altParcaMinBaslangic = yeniAltParcaMinBaslangic;
    sonIsEmriSonuclari = isEmriSonuclari;
  }

  return { isEmriSonuclari: sonIsEmriSonuclari };
}
