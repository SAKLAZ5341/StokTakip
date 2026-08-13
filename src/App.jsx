import React, { useState, useEffect, useMemo, useRef } from "react";
import { Plus, Trash2, ClipboardList, Users, Cog, BarChart3, Factory, X, Lock, Upload, Download, Search, Boxes, FileDown, ChevronDown, ChevronRight, Menu as MenuIcon, UserPlus, Mail, Chrome, Ruler, RefreshCw, Copy, Building2, Bell, ArrowLeft, Home, AlertTriangle, HelpCircle } from "lucide-react";
import { db, auth, digerKullaniciOlustur, eskiMetalErpDb } from "./firebase";
import {
  collection, onSnapshot, addDoc, deleteDoc, doc, updateDoc, writeBatch, query, where, getDocs, getDoc, increment,
} from "firebase/firestore";
import {
  onAuthStateChanged, signInWithEmailAndPassword, signOut, sendPasswordResetEmail,
  GoogleAuthProvider, signInWithPopup, createUserWithEmailAndPassword,
} from "firebase/auth";
import * as XLSX from "xlsx";

const todayISO = () => new Date().toISOString().slice(0, 10);

const MENU = [
  { id: "ana-sayfa", label: "Ana Sayfa", icon: Home },
  {
    id: "hammadde", label: "Hammadde", icon: Factory,
    children: [
      { id: "hammadde-kayit", label: "Hammadde Kaydı" },
      { id: "hammadde-raporu", label: "Hammadde Raporu" },
      { id: "hammadde-sil", label: "Kayıtları Sil" },
    ],
  },
  {
    id: "metal", label: "Metal Ölçü", icon: Ruler,
    children: [
      { id: "metal-hizli", label: "Hızlı KG Hesabı" },
      { id: "metal-gecmis", label: "Geçmiş Ölçümler" },
      { id: "metal-raporu", label: "Metal Ölçü Raporu" },
      { id: "metal-malzeme", label: "Malzeme Tanımları" },
    ],
  },
  {
    id: "depo", label: "Depo Stok", icon: Boxes,
    children: [
      { id: "depo-kart", label: "Stok Kartı Oluştur" },
      { id: "depo-giris", label: "Depo Giriş" },
      { id: "depo-cikis", label: "Depo Çıkış" },
      { id: "depo-hareketler", label: "Stok Hareketleri" },
      { id: "depo-raporu", label: "Depo Stok Raporu" },
      { id: "depo-sil", label: "Stok Kartı Sil" },
    ],
  },
  {
    id: "fason", label: "Fason Takip", icon: Building2,
    children: [
      { id: "fason-ozet", label: "Özet" },
      { id: "fason-firmalar", label: "Firmalar" },
      { id: "fason-isler", label: "İşler" },
      { id: "fason-hareketler", label: "Hareketler" },
      { id: "fason-hatirlaticilar", label: "Hatırlatıcılar" },
      { id: "fason-raporu", label: "Fason Takip Raporu" },
    ],
  },
  { id: "takimlar", label: "Takımlar", icon: Users },
  { id: "makineler", label: "Makineler", icon: Cog },
  { id: "kullanicilar", label: "Kullanıcılar", icon: UserPlus },
  { id: "yardim", label: "Yardım", icon: HelpCircle },
];

// ---------- Excel yardımcıları ----------
function dosyaOku(dosya) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(dosya);
  });
}

// Türkçe sayı formatını ayrıştırır: "1.064,00" -> 1064, "29,00" -> 29, 42 -> 42
function sayiAyristir(v) {
  if (typeof v === "number") return v;
  if (v === undefined || v === null) return 0;
  const s = String(v).trim();
  if (!s) return 0;
  const temiz = s.replace(/\./g, "").replace(",", ".");
  const sayi = parseFloat(temiz);
  return isNaN(sayi) ? 0 : sayi;
}

async function excelDenKayitOku(dosya) {
  const rows = await dosyaOku(dosya);
  if (!rows.length) return [];
  let baslangic = 0;
  const ilkSatir = (rows[0] || []).map((v) => String(v || "").trim().toLowerCase());
  if (ilkSatir[0] && ilkSatir[0].includes("tarih")) baslangic = 1;
  const kayitlar = [];
  for (let i = baslangic; i < rows.length; i++) {
    const r = rows[i] || [];
    const takim = String(r[1] || "").trim();
    const makine = String(r[3] || "").trim();
    if (!takim && !makine) continue;
    let tarih;
    if (r[0] instanceof Date) {
      tarih = r[0].toISOString().slice(0, 10);
    } else {
      const t = String(r[0] || "").trim();
      tarih = t || todayISO();
    }
    kayitlar.push({
      tarih,
      takim,
      magaza: String(r[2] || "").trim(),
      makine,
      urun: String(r[4] || "").trim(),
      adet: sayiAyristir(r[5]),
    });
  }
  return kayitlar;
}

async function excelDenIsimOku(dosya) {
  const rows = await dosyaOku(dosya);
  const baslikKelimeler = ["takım", "takim", "makine", "ad", "isim", "i̇sim", "name"];
  const isimler = rows
    .flat()
    .map((v) => (v === undefined || v === null ? "" : String(v).trim()))
    .filter((v) => v && !baslikKelimeler.includes(v.toLowerCase()));
  return [...new Set(isimler)];
}

async function excelDenDepoOku(dosya) {
  const rows = await dosyaOku(dosya);
  if (!rows.length) return [];
  const normalize = (s) => String(s || "").replace(/İ/g, "I").toLowerCase().trim();
  const ilkSatir = (rows[0] || []).map(normalize);
  const bul = (...anahtarlar) => ilkSatir.findIndex((h) => anahtarlar.some((a) => h.includes(a)));

  const hasHeader = ilkSatir.some((h) => h.includes("stok kod") || h.includes("stok ismi") || h.includes("stok ad"));
  let sutun;
  if (hasHeader) {
    sutun = {
      stokKodu: bul("stok kod"),
      stokAdi: bul("stok ismi", "stok ad"),
      birim: bul("birim"),
      anaGrupKodu: bul("ana grup kod"),
      anaGrupAdi: bul("ana grup ismi", "ana grup ad"),
      altGrupKodu: bul("alt grup kod"),
      altGrupAdi: bul("alt grup ismi", "alt grup ad"),
      miktar: bul("miktar"),
    };
  } else {
    // Başlık satırı yoksa (ör. önceden hazırlanmış eski şablon), sabit eski sırayı kullan
    sutun = { stokKodu: 0, stokAdi: 1, birim: 2, anaGrupKodu: 3, anaGrupAdi: 4, altGrupKodu: 5, altGrupAdi: 6, miktar: 7 };
  }
  const baslangic = hasHeader ? 1 : 0;
  const al = (r, idx) => (idx >= 0 && idx !== undefined ? String(r[idx] || "").trim() : "");

  const kayitlar = [];
  for (let i = baslangic; i < rows.length; i++) {
    const r = rows[i] || [];
    const stokKodu = al(r, sutun.stokKodu);
    const stokAdi = al(r, sutun.stokAdi);
    if (!stokKodu && !stokAdi) continue;
    kayitlar.push({
      stokKodu,
      stokAdi,
      birim: al(r, sutun.birim) || "Adet",
      anaGrupKodu: al(r, sutun.anaGrupKodu),
      anaGrupAdi: al(r, sutun.anaGrupAdi),
      altGrupKodu: al(r, sutun.altGrupKodu),
      altGrupAdi: al(r, sutun.altGrupAdi),
      miktar: sutun.miktar >= 0 ? (sayiAyristir(r[sutun.miktar]) || 0) : 0,
    });
  }
  return kayitlar;
}

// ---------- Metal Talep - kesit hesaplama ----------
const KESIT_TIPLERI = [
  { id: "mil", label: "Mil (Yuvarlak)" },
  { id: "kare", label: "Kare" },
  { id: "lama", label: "Lama" },
  { id: "boru", label: "Boru" },
];
const KESIT_ETIKET = { mil: "Mil", kare: "Kare", lama: "Lama", boru: "Boru" };
const OLCU_ALANLARI = {
  mil: [{ id: "cap", label: "Çap (mm)", def: 20 }],
  kare: [{ id: "kenar", label: "Kenar (mm)", def: 20 }],
  lama: [{ id: "kalinlik", label: "Kalınlık (mm)", def: 10 }, { id: "genislik", label: "Genişlik (mm)", def: 50 }],
  boru: [{ id: "disCap", label: "Dış Çap (mm)", def: 33.7 }, { id: "etKalinligi", label: "Et Kalınlığı (mm)", def: 3.2 }],
};
function alanMm2(tur, dims) {
  if (tur === "mil") return (Math.PI / 4) * Math.pow(dims.cap || 0, 2);
  if (tur === "kare") return Math.pow(dims.kenar || 0, 2);
  if (tur === "lama") return (dims.kalinlik || 0) * (dims.genislik || 0);
  if (tur === "boru") { const s = dims.etKalinligi || 0, D = dims.disCap || 0; return Math.PI * s * (D - s); }
  return 0;
}
function kgMetre(tur, dims, yogunluk) { return alanMm2(tur, dims) * (yogunluk || 0) * 0.001; }
function olcuEtiketi(tur, dims) {
  if (tur === "mil") return `Ø${dims.cap} mm`;
  if (tur === "kare") return `${dims.kenar}x${dims.kenar} mm`;
  if (tur === "lama") return `${dims.kalinlik}x${dims.genislik} mm`;
  if (tur === "boru") return `Ø${dims.disCap} x ${dims.etKalinligi} mm`;
  return "";
}
function boyMetreCevir(deger, birim) {
  const v = Number(deger) || 0;
  if (birim === "mm") return v / 1000;
  if (birim === "cm") return v / 100;
  return v;
}
const VARSAYILAN_MALZEMELER = [
  { ad: "Çelik", yogunluk: 7.85 },
  { ad: "Paslanmaz", yogunluk: 7.90 },
  { ad: "Bronz", yogunluk: 8.80 },
  { ad: "Kestamid", yogunluk: 1.15 },
  { ad: "Alüminyum", yogunluk: 2.70 },
];

// Malzeme seçim ekranlarında her zaman en az bu 5 malzeme hazır dursun -
// "Malzeme Tanımları" ekranından ekstra malzeme eklerse (ya da yoğunluğunu
// değiştirirse) o da bu listeye dahil olur / öncelik alır.
function birlesikMalzemeler(metalMalzemeler) {
  const ozelAdlar = new Set(metalMalzemeler.map((m) => m.ad.toLowerCase()));
  const varsayilanlar = VARSAYILAN_MALZEMELER
    .filter((m) => !ozelAdlar.has(m.ad.toLowerCase()))
    .map((m) => ({ id: "varsayilan-" + m.ad, ...m }));
  return [...varsayilanlar, ...metalMalzemeler];
}

async function excelDenMetalOku(dosya, metalMalzemeler) {
  const rows = await dosyaOku(dosya);
  if (!rows.length) return [];
  let baslangic = 0;
  const ilkSatir = (rows[0] || []).map((v) => String(v || "").trim().toLowerCase());
  if (ilkSatir[0] && ilkSatir[0].includes("talep")) baslangic = 1;
  const turMap = { mil: "mil", kare: "kare", lama: "lama", boru: "boru" };
  const kayitlar = [];
  for (let i = baslangic; i < rows.length; i++) {
    const r = rows[i] || [];
    const malzemeAdi = String(r[2] || "").trim();
    if (!malzemeAdi) continue;
    const turRaw = String(r[1] || "").trim().toLowerCase();
    const tur = turMap[turRaw] || "mil";
    const malzeme = metalMalzemeler.find((m) => m.ad.toLowerCase() === malzemeAdi.toLowerCase());
    const yogunluk = malzeme ? malzeme.yogunluk : 7.85;
    const olcu1 = sayiAyristir(r[3]);
    const olcu2 = sayiAyristir(r[4]);
    const dims = tur === "mil" ? { cap: olcu1 } : tur === "kare" ? { kenar: olcu1 } : tur === "lama" ? { kalinlik: olcu1, genislik: olcu2 } : { disCap: olcu1, etKalinligi: olcu2 };
    const boy = sayiAyristir(r[5]);
    const adet = sayiAyristir(r[6]) || 1;
    const fiyat = sayiAyristir(r[7]);
    const birimKg = kgMetre(tur, dims, yogunluk) * boy;
    const toplamKg = birimKg * adet;
    kayitlar.push({
      talepNo: String(r[0] || "").trim() || "—",
      tur, malzemeAdi, yogunluk, dims, dimLabel: olcuEtiketi(tur, dims),
      boy, adet, birimKg, toplamKg,
      fiyat: fiyat || null, tutar: fiyat ? toplamKg * fiyat : 0,
    });
  }
  return kayitlar;
}

// ---------- Fason Takip - yardımcılar ----------
function paraTR(n) { return (Number(n) || 0).toLocaleString("tr-TR", { style: "currency", currency: "TRY", maximumFractionDigits: 2 }); }
const FASON_DURUM = {
  bekliyor: { label: "Bekliyor", renk: "#e8a33d" },
  uretimde: { label: "Üretimde", renk: "#2dd4bf" },
  tamamlandi: { label: "Tamamlandı", renk: "#4b8f5e" },
};
const FASON_KALITE = {
  okeylendi: { label: "Okeylendi", renk: "#4b8f5e" },
  red: { label: "Red", renk: "#e07a6b" },
  olcumde: { label: "Ölçümde", renk: "#e8a33d" },
};

async function excelDenFasonFirmaOku(dosya) {
  const rows = await dosyaOku(dosya);
  if (!rows.length) return [];
  let baslangic = 0;
  const ilkSatir = (rows[0] || []).map((v) => String(v || "").trim().toLowerCase());
  if (ilkSatir[0] && ilkSatir[0].includes("firma")) baslangic = 1;
  const kayitlar = [];
  for (let i = baslangic; i < rows.length; i++) {
    const r = rows[i] || [];
    const ad = String(r[0] || "").trim();
    if (!ad) continue;
    kayitlar.push({ ad, yetkili: String(r[1] || "").trim(), not: String(r[2] || "").trim() });
  }
  return kayitlar;
}

async function excelDenFasonIsOku(dosya, fasonFirmalar) {
  const rows = await dosyaOku(dosya);
  if (!rows.length) return [];
  let baslangic = 0;
  const ilkSatir = (rows[0] || []).map((v) => String(v || "").trim().toLowerCase());
  if (ilkSatir[0] && ilkSatir[0].includes("firma")) baslangic = 1;
  const durumMap = { bekliyor: "bekliyor", "üretimde": "uretimde", uretimde: "uretimde", tamamlandı: "tamamlandi", tamamlandi: "tamamlandi" };
  const kayitlar = [];
  for (let i = baslangic; i < rows.length; i++) {
    const r = rows[i] || [];
    const firmaAdi = String(r[0] || "").trim();
    const projeAdi = String(r[2] || "").trim();
    if (!firmaAdi || !projeAdi) continue;
    const firma = fasonFirmalar.find((f) => f.ad.toLowerCase() === firmaAdi.toLowerCase());
    kayitlar.push({
      firmaId: firma ? firma.id : null, firmaAdiGecici: firma ? null : firmaAdi,
      projeKodu: String(r[1] || "").trim(), projeAdi,
      miktar: String(r[3] || "").trim(), ucret: sayiAyristir(r[4]),
      resimRef: String(r[5] || "").trim(), aciklama: String(r[6] || "").trim(),
      durum: durumMap[String(r[7] || "").trim().toLowerCase()] || "bekliyor",
      olusturmaTarihi: String(r[8] || "").trim() || todayISO(),
    });
  }
  return kayitlar;
}

async function excelDenFasonHareketOku(dosya, fasonFirmalar, fasonIsler) {
  const rows = await dosyaOku(dosya);
  if (!rows.length) return [];
  let baslangic = 0;
  const ilkSatir = (rows[0] || []).map((v) => String(v || "").trim().toLowerCase());
  if (ilkSatir[0] && ilkSatir[0].includes("firma")) baslangic = 1;
  const kayitlar = [];
  for (let i = baslangic; i < rows.length; i++) {
    const r = rows[i] || [];
    const firmaAdi = String(r[0] || "").trim();
    const projeKodu = String(r[1] || "").trim();
    const urunAdi = String(r[4] || "").trim();
    if (!firmaAdi || !urunAdi) continue;
    const firma = fasonFirmalar.find((f) => f.ad.toLowerCase() === firmaAdi.toLowerCase());
    const is = firma ? fasonIsler.find((j) => j.firmaId === firma.id && (j.projeKodu || "").toLowerCase() === projeKodu.toLowerCase()) : null;
    if (!is) continue;
    const tipRaw = String(r[3] || "").trim().toLowerCase();
    const tip = tipRaw.startsWith("gelen") ? "gelen" : "giden";
    kayitlar.push({
      isId: is.id, tip, urunAdi,
      malzemeCinsi: String(r[5] || "").trim(), kalite: String(r[6] || "").trim(), aciklama: String(r[7] || "").trim(),
      miktar: sayiAyristir(r[8]), birim: String(r[9] || "").trim(),
      birimFiyat: sayiAyristir(r[10]), tarih: String(r[11] || "").trim() || todayISO(),
      not: String(r[12] || "").trim(),
    });
  }
  return kayitlar;
}


async function excelDenHammaddeOku(dosya) {
  const rows = await dosyaOku(dosya);
  if (!rows.length) return [];
  let baslangic = 0;
  const ilkSatir = (rows[0] || []).map((v) => String(v || "").trim().toLowerCase());
  if (ilkSatir[0] && (ilkSatir[0].includes("cari") || ilkSatir[0].includes("firma"))) {
    baslangic = 1;
  }
  const kayitlar = [];
  for (let i = baslangic; i < rows.length; i++) {
    const r = rows[i] || [];
    const cari = String(r[0] || "").trim();
    if (!cari) continue;
    kayitlar.push({
      cari,
      projeKodu: String(r[1] || "").trim(),
      projeAdi: String(r[2] || "").trim(),
      kalite: String(r[3] || "").trim(),
      aciklama1: String(r[4] || "").trim(),
      aciklama2: String(r[5] || "").trim(),
      miktar: sayiAyristir(r[6]),
      durumu: String(r[7] || "").trim(),
      tamamlandi: false,
    });
  }
  return kayitlar;
}

function excelIndir(veri, dosyaAdi, sayfaAdi) {
  const ws = XLSX.utils.json_to_sheet(veri);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sayfaAdi);
  XLSX.writeFile(wb, dosyaAdi);
}

function sablonIndir(basliklar, ornekSatirlar, dosyaAdi, sayfaAdi) {
  const ws = XLSX.utils.aoa_to_sheet([basliklar, ...ornekSatirlar]);
  ws["!cols"] = basliklar.map(() => ({ wch: 22 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sayfaAdi);
  XLSX.writeFile(wb, dosyaAdi);
}

// Firestore'a güvenli toplu yazma: küçük parçalar halinde gönderir, bir parça
// başarısız olursa (bağlantı sorunu vb.) 3 kere tekrar dener, atlamadan devam
// eder ve en sonunda kaç kaydın başarılı/başarısız olduğunu döner.
async function guvenliTopluYaz(koleksiyon, kayitlar, ilerlemeCB) {
  const PARCA = 50;
  let basarili = 0;
  let basarisiz = 0;
  for (let i = 0; i < kayitlar.length; i += PARCA) {
    const dilim = kayitlar.slice(i, i + PARCA);
    let deneme = 0;
    let tamam = false;
    while (deneme < 3 && !tamam) {
      try {
        const batch = writeBatch(db);
        dilim.forEach((k) => {
          const ref = doc(collection(db, koleksiyon));
          batch.set(ref, k);
        });
        await batch.commit();
        tamam = true;
      } catch (err) {
        deneme++;
        console.error(`${koleksiyon} yazma hatası (deneme ${deneme}/3):`, err);
        if (deneme < 3) await new Promise((r) => setTimeout(r, 700 * deneme));
      }
    }
    if (tamam) basarili += dilim.length; else basarisiz += dilim.length;
    if (ilerlemeCB) ilerlemeCB(basarili + basarisiz, kayitlar.length, basarisiz);
  }
  return { basarili, basarisiz };
}

// ---------- Şifre Kapısı ----------
function GirisEkrani() {
  const [email, setEmail] = useState("");
  const [sifre, setSifre] = useState("");
  const [hata, setHata] = useState("");
  const [gonderiliyor, setGonderiliyor] = useState(false);
  const [sifremiUnuttum, setSifremiUnuttum] = useState(false);
  const [sifirlamaMesaji, setSifirlamaMesaji] = useState("");
  const [kayitModu, setKayitModu] = useState(false);
  const [ad, setAd] = useState("");

  const dene = async () => {
    if (!email || !sifre) {
      setHata("E-posta ve şifre gerekli.");
      return;
    }
    setGonderiliyor(true);
    setHata("");
    try {
      await signInWithEmailAndPassword(auth, email.trim(), sifre);
    } catch (err) {
      const kod = err?.code || "";
      if (kod.includes("invalid-credential") || kod.includes("wrong-password") || kod.includes("user-not-found")) {
        setHata("E-posta veya şifre hatalı.");
      } else if (kod.includes("too-many-requests")) {
        setHata("Çok fazla hatalı deneme. Biraz sonra tekrar dene.");
      } else {
        setHata("Giriş yapılamadı: " + kod);
      }
    }
    setGonderiliyor(false);
  };

  const kayitOl = async () => {
    if (!email || !sifre) {
      setHata("E-posta ve şifre gerekli.");
      return;
    }
    if (sifre.length < 6) {
      setHata("Şifre en az 6 karakter olmalı.");
      return;
    }
    setGonderiliyor(true);
    setHata("");
    try {
      const sonuc = await createUserWithEmailAndPassword(auth, email.trim(), sifre);
      await addDoc(collection(db, "kullanicilar"), {
        ad: ad.trim(), email: email.trim(), emailKucuk: email.trim().toLowerCase(),
        tur: "sifreli", eklenmeTarihi: Date.now(),
      });
    } catch (err) {
      const kod = err?.code || "";
      if (kod.includes("email-already-in-use")) setHata("Bu e-posta zaten kayıtlı, giriş yapmayı dene.");
      else if (kod.includes("invalid-email")) setHata("E-posta adresi geçersiz.");
      else if (kod.includes("weak-password")) setHata("Şifre çok zayıf, en az 6 karakter olmalı.");
      else setHata("Kayıt olunamadı: " + kod);
    }
    setGonderiliyor(false);
  };

  const sifreSifirla = async () => {
    if (!email) {
      setHata("Önce e-posta adresini yaz.");
      return;
    }
    setGonderiliyor(true);
    setHata("");
    try {
      await sendPasswordResetEmail(auth, email.trim());
      setSifirlamaMesaji("Şifre sıfırlama linki e-postana gönderildi.");
    } catch (err) {
      setHata("Gönderilemedi: " + (err?.code || "bilinmeyen hata"));
    }
    setGonderiliyor(false);
  };

  const googleIleGiris = async () => {
    setGonderiliyor(true);
    setHata("");
    try {
      const provider = new GoogleAuthProvider();
      await signInWithPopup(auth, provider);
      // İzin kontrolü artık App bileşeninde merkezi olarak yapılıyor,
      // burada tekrar kontrol etmiyoruz (sayfa geçişinde hatayı kaybetmemek için).
    } catch (err) {
      const kod = err?.code || "";
      if (!kod.includes("popup-closed-by-user")) {
        setHata("Google girişi başarısız: " + kod);
      }
    }
    setGonderiliyor(false);
  };

  return (
    <div style={{ minHeight: "100vh", background: "#142a30", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', system-ui, sans-serif" }}>
      <div style={{ background: "#1b333c", border: "1px solid #2a4b52", borderRadius: 12, padding: 32, width: 320 }}>
        <div style={{ width: 40, height: 40, borderRadius: 10, background: "#2dd4bf", display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 16 }}>
          <Lock size={20} color="#142a30" />
        </div>
        <div style={{ color: "#e7e5e0", fontWeight: 700, fontSize: 17, marginBottom: 2, letterSpacing: -0.2 }}>SAKLAZ</div>
        <div style={{ color: "#2dd4bf", fontWeight: 600, fontSize: 11.5, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 14 }}>Üretim ERP</div>
        <div style={{ color: "#8b929a", fontSize: 12.5, marginBottom: 18 }}>
          {sifremiUnuttum ? "Şifre sıfırlama linki için e-postanı gir" : kayitModu ? "Yeni hesap oluştur" : "Devam etmek için giriş yap"}
        </div>
        {kayitModu && !sifremiUnuttum && (
          <input
            type="text"
            value={ad}
            onChange={(e) => setAd(e.target.value)}
            placeholder="Ad Soyad (opsiyonel)"
            style={{ width: "100%", background: "#142a30", border: "1px solid #3d6169", borderRadius: 7, padding: "10px 12px", color: "#e7e5e0", fontSize: 14, outline: "none", marginBottom: 10 }}
          />
        )}
        <input
          type="email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && !sifremiUnuttum && (kayitModu ? kayitOl() : dene())}
          placeholder="E-posta"
          style={{ width: "100%", background: "#142a30", border: `1px solid ${hata ? "#c0392b" : "#3d6169"}`, borderRadius: 7, padding: "10px 12px", color: "#e7e5e0", fontSize: 14, outline: "none", marginBottom: 10 }}
        />
        {!sifremiUnuttum && (
          <input
            type="password"
            value={sifre}
            onChange={(e) => setSifre(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && (kayitModu ? kayitOl() : dene())}
            placeholder={kayitModu ? "Şifre (en az 6 karakter)" : "Şifre"}
            style={{ width: "100%", background: "#142a30", border: `1px solid ${hata ? "#c0392b" : "#3d6169"}`, borderRadius: 7, padding: "10px 12px", color: "#e7e5e0", fontSize: 14, outline: "none", marginBottom: 12 }}
          />
        )}
        <button
          onClick={sifremiUnuttum ? sifreSifirla : kayitModu ? kayitOl : dene}
          disabled={gonderiliyor}
          style={{ width: "100%", background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 0", fontWeight: 700, fontSize: 13.5, cursor: "pointer", marginTop: sifremiUnuttum ? 12 : 0 }}
        >
          {gonderiliyor ? "Gönderiliyor…" : sifremiUnuttum ? "Sıfırlama Linki Gönder" : kayitModu ? "Kayıt Ol" : "Giriş Yap"}
        </button>
        {hata && <div style={{ color: "#e07a6b", fontSize: 12, marginTop: 10 }}>{hata}</div>}
        {sifirlamaMesaji && <div style={{ color: "#2dd4bf", fontSize: 12, marginTop: 10 }}>{sifirlamaMesaji}</div>}
        {!sifremiUnuttum && !kayitModu && (
          <>
            <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "16px 0" }}>
              <div style={{ flex: 1, height: 1, background: "#2a4b52" }} />
              <span style={{ color: "#6b7178", fontSize: 11 }}>veya</span>
              <div style={{ flex: 1, height: 1, background: "#2a4b52" }} />
            </div>
            <button
              onClick={googleIleGiris}
              disabled={gonderiliyor}
              style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, background: "#142a30", color: "#e7e5e0", border: "1px solid #3d6169", borderRadius: 7, padding: "10px 0", fontWeight: 600, fontSize: 13.5, cursor: "pointer" }}
            >
              <Chrome size={16} /> Google ile Giriş Yap
            </button>
          </>
        )}
        {!sifremiUnuttum && (
          <button
            onClick={() => { setKayitModu((s) => !s); setHata(""); }}
            style={{ width: "100%", background: "none", border: "none", color: "#8b929a", fontSize: 12, cursor: "pointer", marginTop: 14, textDecoration: "underline" }}
          >
            {kayitModu ? "Zaten hesabım var, girişe dön" : "Hesabın yok mu? Kayıt Ol"}
          </button>
        )}
        {!kayitModu && (
          <button
            onClick={() => { setSifremiUnuttum((s) => !s); setHata(""); setSifirlamaMesaji(""); }}
            style={{ width: "100%", background: "none", border: "none", color: "#8b929a", fontSize: 12, cursor: "pointer", marginTop: 8, textDecoration: "underline" }}
          >
            {sifremiUnuttum ? "Girişe geri dön" : "Şifremi unuttum"}
          </button>
        )}
      </div>
    </div>
  );
}

export default function App() {
  const [kullanici, setKullanici] = useState(undefined); // undefined: yükleniyor, null: girmemiş, object: girmiş

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (u) => {
      if (u) {
        // İlk girişte otomatik olarak "kullanicilar" listesine kaydını oluştur
        // (Google isim varsa otomatik doldurur, yoksa Ana Sayfa'dan sonra elle girilebilir).
        try {
          const emailKucuk = (u.email || "").toLowerCase();
          const sorgu = query(collection(db, "kullanicilar"), where("emailKucuk", "==", emailKucuk));
          const sonuc = await getDocs(sorgu);
          if (sonuc.empty) {
            const googleIle = u.providerData.some((p) => p.providerId === "google.com");
            await addDoc(collection(db, "kullanicilar"), {
              ad: u.displayName || "", email: u.email, emailKucuk,
              tur: googleIle ? "google" : "sifreli", eklenmeTarihi: Date.now(),
            });
          }
        } catch (err) {
          console.error(err);
        }
      }
      setKullanici(u);
    });
    return unsub;
  }, []);

  if (kullanici === undefined) {
    return (
      <div style={{ minHeight: "100vh", background: "#142a30", display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ color: "#2dd4bf", fontFamily: "monospace", fontSize: 14, letterSpacing: 1 }}>YÜKLENİYOR…</div>
      </div>
    );
  }
  if (!kullanici) return <GirisEkrani />;
  return <Panel onCikis={() => signOut(auth)} kullanici={kullanici} />;
}

function Panel({ onCikis, kullanici }) {
  const [tab, setTab] = useState("ana-sayfa");
  const [acikGruplar, setAcikGruplar] = useState(new Set(["uretim"]));
  const [mobilMenuAcik, setMobilMenuAcik] = useState(false);

  // Tarayıcının geri/ileri tuşları uygulama içinde menüler arasında gezinsin,
  // uygulamadan çıkıp giriş ekranına dönmesin.
  useEffect(() => {
    if (!window.history.state || !window.history.state.uretimTakipTab) {
      window.history.replaceState({ uretimTakipTab: "ana-sayfa" }, "");
    } else {
      setTab(window.history.state.uretimTakipTab);
    }
    const onPopState = (e) => {
      if (e.state && e.state.uretimTakipTab) setTab(e.state.uretimTakipTab);
    };
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, []);

  const [teams, setTeams] = useState([]);
  const [machines, setMachines] = useState([]);
  const [records, setRecords] = useState([]);
  const [hammaddeler, setHammaddeler] = useState([]);
  const [depoStok, setDepoStok] = useState([]);
  const [depoHareketler, setDepoHareketler] = useState([]);
  const [metalMalzemeler, setMetalMalzemeler] = useState([]);
  const [metalTalepler, setMetalTalepler] = useState([]);
  const [fasonFirmalar, setFasonFirmalar] = useState([]);
  const [fasonIsler, setFasonIsler] = useState([]);
  const [fasonHareketler, setFasonHareketler] = useState([]);
  const [fasonHatirlaticilar, setFasonHatirlaticilar] = useState([]);
  const [kullanicilar, setKullanicilar] = useState([]);

  // Firestore canlı dinleme - herkes aynı anda güncel veriyi görür
  useEffect(() => {
    const unsub1 = onSnapshot(collection(db, "teams"), (snap) =>
      setTeams(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub2 = onSnapshot(collection(db, "machines"), (snap) =>
      setMachines(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub3 = onSnapshot(collection(db, "records"), (snap) =>
      setRecords(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub4 = onSnapshot(collection(db, "hammadde"), (snap) =>
      setHammaddeler(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub5 = onSnapshot(collection(db, "depo_stok"), (snap) =>
      setDepoStok(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub6 = onSnapshot(collection(db, "depo_hareketler"), (snap) =>
      setDepoHareketler(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub7 = onSnapshot(collection(db, "metal_malzemeler"), (snap) =>
      setMetalMalzemeler(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub8 = onSnapshot(collection(db, "metal_talepler"), (snap) =>
      setMetalTalepler(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub9 = onSnapshot(collection(db, "fason_firmalar"), (snap) =>
      setFasonFirmalar(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub10 = onSnapshot(collection(db, "fason_isler"), (snap) =>
      setFasonIsler(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub11 = onSnapshot(collection(db, "fason_hareketler"), (snap) =>
      setFasonHareketler(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub12 = onSnapshot(collection(db, "fason_hatirlaticilar"), (snap) =>
      setFasonHatirlaticilar(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub13 = onSnapshot(collection(db, "kullanicilar"), (snap) =>
      setKullanicilar(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    return () => { unsub1(); unsub2(); unsub3(); unsub4(); unsub5(); unsub6(); unsub7(); unsub8(); unsub9(); unsub10(); unsub11(); unsub12(); unsub13(); };
  }, []);

  const secimYap = (id) => {
    window.history.pushState({ uretimTakipTab: id }, "");
    setTab(id);
    setMobilMenuAcik(false);
  };

  const geriGit = () => window.history.back();

  const grupToggle = (id) => {
    setAcikGruplar((s) => {
      const yeni = new Set(s);
      if (yeni.has(id)) yeni.delete(id); else yeni.add(id);
      return yeni;
    });
  };

  const aktifBaslik = () => {
    if (tab === "ana-sayfa") return "Ana Sayfa";
    if (tab === "stok-kayit") return "Üretim Kaydı";
    if (tab === "hammadde-kayit") return "Hammadde Kaydı";
    if (tab === "metal-hizli") return "Hızlı KG Hesabı";
    if (tab === "metal-gecmis") return "Geçmiş Ölçümler";
    if (tab === "metal-malzeme") return "Malzeme Tanımları";
    if (tab === "depo-kart") return "Stok Kartı Oluştur";
    if (tab === "depo-giris") return "Depo Giriş";
    if (tab === "depo-cikis") return "Depo Çıkış";
    if (tab === "fason-ozet") return "Fason Takip Özeti";
    if (tab === "fason-firmalar") return "Fason Firmalar";
    if (tab === "fason-isler") return "Fason İşler";
    if (tab === "fason-hareketler") return "Fason Hareketleri";
    if (tab === "fason-hatirlaticilar") return "Fason Hatırlatıcıları";
    if (tab === "depo-hareketler") return "Stok Hareketleri";
    if (tab === "stok-raporu") return "Üretim Raporu";
    if (tab === "hammadde-raporu") return "Hammadde Raporu";
    if (tab === "metal-raporu") return "Metal Ölçü Raporu";
    if (tab === "depo-raporu") return "Depo Stok Raporu";
    if (tab === "fason-raporu") return "Fason Takip Raporu";
    if (tab === "stok-sil") return "Üretim Kayıtları Sil";
    if (tab === "hammadde-sil") return "Hammadde Kayıtları Sil";
    if (tab === "depo-sil") return "Depo Stok Sil";
    if (tab === "takimlar") return "Takımlar";
    if (tab === "makineler") return "Makineler";
    if (tab === "kullanicilar") return "Kullanıcılar";
    if (tab === "yardim") return "Yardım";
    return "";
  };

  return (
    <div style={{ minHeight: "100vh", background: "#142a30", color: "#e7e5e0", fontFamily: "'Inter', system-ui, sans-serif", display: "flex", flexDirection: "column" }}>
      <style>{`
        * { box-sizing: border-box; }
        input, select, button { font-family: inherit; }
        ::placeholder { color: #6b7178; }
        .card { background: #1b333c; border: 1px solid #2a4b52; border-radius: 10px; }
        .field-label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.08em; color: #8b929a; margin-bottom: 6px; display: block; font-weight: 600; }
        .input { width: 100%; background: #142a30; border: 1px solid #3d6169; border-radius: 7px; padding: 10px 12px; color: #e7e5e0; font-size: 14px; outline: none; transition: border-color .15s; }
        .input:focus { border-color: #2dd4bf; }
        table { border-collapse: collapse; width: 100%; }
        th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em; color: #8b929a; padding: 10px 12px; border-bottom: 1px solid #2a4b52; font-weight: 600; white-space: nowrap; }
        td { padding: 10px 12px; border-bottom: 1px solid #24424a; font-size: 13.5px; }
        tr:hover td { background: #274852; }
        .btn-ghost { display: flex; align-items: center; gap: 6px; background: transparent; border: 1px solid #3d6169; color: #c7cbd1; border-radius: 7px; padding: 8px 13px; font-size: 12.5px; font-weight: 600; cursor: pointer; white-space: nowrap; }
        .btn-ghost:hover { border-color: #2dd4bf; color: #2dd4bf; }
        .pill { display: inline-block; padding: 3px 9px; border-radius: 20px; font-size: 11px; font-weight: 700; background: #113330; color: #2dd4bf; border: 1px solid #1f4d47; white-space: nowrap; }
        .navbtn { display: flex; align-items: center; gap: 10px; width: 100%; text-align: left; padding: 10px 12px; border-radius: 8px; border: none; cursor: pointer; font-size: 13.5px; font-weight: 600; background: transparent; color: #c7cbd1; }
        .navbtn:hover { background: #274852; }
        .navbtn.active { background: #2dd4bf; color: #142a30; }
        .navsub { display: flex; align-items: center; gap: 8px; width: 100%; text-align: left; padding: 8px 12px 8px 38px; border-radius: 8px; border: none; cursor: pointer; font-size: 13px; font-weight: 600; background: transparent; color: #9aa0a8; }
        .navsub:hover { background: #274852; color: #c7cbd1; }
        .navsub.active { background: #3a3220; color: #2dd4bf; }
        .mobil-menu-btn { display: none; }
        @media (max-width: 820px) {
          .mobil-menu-btn { display: flex; }
          .sidebar { position: fixed; inset: 0 0 0 0; z-index: 40; width: 260px !important; transform: translateX(-100%); transition: transform .2s; }
          .sidebar.open { transform: translateX(0); }
          .sidebar-backdrop { position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 30; }
        }
      `}</style>

      <header style={{ borderBottom: "1px solid #2a4b52", padding: "14px 20px", display: "flex", alignItems: "center", gap: 12 }}>
        <button
          className="btn-ghost mobil-menu-btn"
          onClick={() => setMobilMenuAcik(true)}
        >
          <MenuIcon size={16} />
        </button>
        <button
          onClick={geriGit}
          title="Geri"
          style={{ background: "none", border: "1px solid #3d6169", color: "#c7cbd1", borderRadius: 7, width: 34, height: 34, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 }}
        >
          <ArrowLeft size={16} />
        </button>
        <div style={{ width: 34, height: 34, borderRadius: 8, background: "#2dd4bf", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Factory size={19} color="#142a30" />
        </div>
        <div style={{ flex: 1 }}>
          <div style={{ fontWeight: 700, fontSize: 16, letterSpacing: -0.2 }}>SAKLAZ <span style={{ color: "#2dd4bf" }}>ERP</span></div>
          <div style={{ fontSize: 11.5, color: "#8b929a" }}>{aktifBaslik()}</div>
        </div>
        <button onClick={onCikis} style={{ background: "none", border: "1px solid #3d6169", color: "#8b929a", borderRadius: 7, padding: "7px 12px", fontSize: 12.5, cursor: "pointer", display: "flex", alignItems: "center", gap: 8 }}>
          {kullanici?.email && <span style={{ color: "#6b7178", fontSize: 11.5 }}>{kullanici.email}</span>}
          Çıkış Yap
        </button>
      </header>

      <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
        {mobilMenuAcik && <div className="sidebar-backdrop" onClick={() => setMobilMenuAcik(false)} />}
        <aside className={`sidebar${mobilMenuAcik ? " open" : ""}`} style={{ width: 230, flexShrink: 0, borderRight: "1px solid #2a4b52", background: "#1b333c", padding: "16px 10px", display: "flex", flexDirection: "column", gap: 2 }}>
          {MENU.map((item) => {
            const Icon = item.icon;
            if (item.children) {
              const acik = acikGruplar.has(item.id);
              const grupAktif = item.children.some((c) => c.id === tab);
              return (
                <div key={item.id}>
                  <button
                    className={`navbtn${grupAktif && !acik ? " active" : ""}`}
                    onClick={() => grupToggle(item.id)}
                  >
                    <Icon size={16} />
                    <span style={{ flex: 1 }}>{item.label}</span>
                    {acik ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                  </button>
                  {acik && item.children.map((c) => (
                    <button
                      key={c.id}
                      className={`navsub${tab === c.id ? " active" : ""}`}
                      onClick={() => secimYap(c.id)}
                    >
                      {c.label}
                    </button>
                  ))}
                </div>
              );
            }
            return (
              <button
                key={item.id}
                className={`navbtn${tab === item.id ? " active" : ""}`}
                onClick={() => secimYap(item.id)}
              >
                <Icon size={16} />
                <span>{item.label}</span>
              </button>
            );
          })}
        </aside>

        <main style={{ flex: 1, padding: 24, overflowY: "auto", minWidth: 0 }}>
          <div style={{ maxWidth: 1200, margin: "0 auto" }}>
            {tab === "ana-sayfa" && <AnaSayfa
              kullanici={kullanici} git={secimYap} kullanicilar={kullanicilar}
              teams={teams} machines={machines} records={records}
              hammaddeler={hammaddeler} depoStok={depoStok} depoHareketler={depoHareketler}
              metalTalepler={metalTalepler}
              fasonFirmalar={fasonFirmalar} fasonIsler={fasonIsler} fasonHareketler={fasonHareketler} fasonHatirlaticilar={fasonHatirlaticilar}
            />}
            {tab === "stok-kayit" && <KayitEkle teams={teams} machines={machines} records={records} />}
            {tab === "hammadde-kayit" && <HammaddeTakip hammaddeler={hammaddeler} />}
            {tab === "metal-hizli" && <MetalHizliHesap metalMalzemeler={metalMalzemeler} kullanici={kullanici} />}
            {tab === "metal-gecmis" && <MetalGecmisOlcumler metalTalepler={metalTalepler} metalMalzemeler={metalMalzemeler} />}
            {tab === "metal-malzeme" && <MetalMalzemeYonetimi metalMalzemeler={metalMalzemeler} />}
            {tab === "depo-kart" && <DepoStokKart depoStok={depoStok} kullanici={kullanici} />}
            {tab === "depo-giris" && <DepoGiris depoStok={depoStok} kullanici={kullanici} depoHareketler={depoHareketler} />}
            {tab === "depo-cikis" && <DepoStokCikis depoStok={depoStok} machines={machines} kullanici={kullanici} depoHareketler={depoHareketler} />}
            {tab === "fason-ozet" && <FasonOzet fasonFirmalar={fasonFirmalar} fasonIsler={fasonIsler} fasonHareketler={fasonHareketler} fasonHatirlaticilar={fasonHatirlaticilar} />}
            {tab === "fason-firmalar" && <FasonFirmalar fasonFirmalar={fasonFirmalar} fasonIsler={fasonIsler} fasonHareketler={fasonHareketler} />}
            {tab === "fason-isler" && <FasonIsler fasonFirmalar={fasonFirmalar} fasonIsler={fasonIsler} fasonHareketler={fasonHareketler} />}
            {tab === "fason-hareketler" && <FasonHareketler fasonFirmalar={fasonFirmalar} fasonIsler={fasonIsler} fasonHareketler={fasonHareketler} />}
            {tab === "fason-hatirlaticilar" && <FasonHatirlaticilar fasonIsler={fasonIsler} fasonHatirlaticilar={fasonHatirlaticilar} />}
            {tab === "depo-hareketler" && <DepoHareketleri depoHareketler={depoHareketler} />}
            {tab === "stok-raporu" && <UretimRaporu teams={teams} machines={machines} records={records} />}
            {tab === "hammadde-raporu" && <HammaddeRaporlari hammaddeler={hammaddeler} />}
            {tab === "metal-raporu" && <MetalOlcuRaporu metalTalepler={metalTalepler} metalMalzemeler={metalMalzemeler} />}
            {tab === "depo-raporu" && <DepoStokRaporu depoStok={depoStok} depoHareketler={depoHareketler} />}
            {tab === "fason-raporu" && <FasonTakipRaporu fasonFirmalar={fasonFirmalar} fasonIsler={fasonIsler} fasonHareketler={fasonHareketler} />}
            {tab === "stok-sil" && <StokSilme records={records} />}
            {tab === "hammadde-sil" && <HammaddeSilme hammaddeler={hammaddeler} />}
            {tab === "depo-sil" && <DepoSilme depoStok={depoStok} />}
            {tab === "takimlar" && <ListeYonetimi title="Takım" baslikCogul="Takımlar" koleksiyon="teams" placeholder="Örn: Kesim Takım 1" items={teams} icon={Users} />}
            {tab === "makineler" && <ListeYonetimi title="Makine" baslikCogul="Makineler" koleksiyon="machines" placeholder="Makine listesini buradan ekleyin" items={machines} icon={Cog} />}
            {tab === "kullanicilar" && <KullaniciYonetimi mevcutKullanici={kullanici} />}
            {tab === "yardim" && <YardimEkrani git={secimYap} />}
          </div>
        </main>
      </div>
    </div>
  );
}

// ---------- Ana Sayfa (Kontrol Paneli) ----------
function AnaSayfa({ kullanici, git, kullanicilar, teams, machines, records, hammaddeler, depoStok, depoHareketler, metalTalepler, fasonFirmalar, fasonIsler, fasonHareketler, fasonHatirlaticilar }) {
  const bugun = todayISO();
  const buAy = bugun.slice(0, 7);

  const emailKucuk = (kullanici?.email || "").toLowerCase();
  const benimKaydim = kullanicilar.find((k) => (k.emailKucuk || (k.email || "").toLowerCase()) === emailKucuk);
  const gosterilenIsim = benimKaydim?.ad || kullanici?.displayName || (kullanici?.email || "").split("@")[0];

  const [isimDuzenle, setIsimDuzenle] = useState(false);
  const [isimGirdi, setIsimGirdi] = useState(benimKaydim?.ad || "");
  const [isimMsg, setIsimMsg] = useState("");

  const isimKaydet = async () => {
    const yeniAd = isimGirdi.trim();
    if (!yeniAd) { setIsimMsg("İsim boş olamaz."); setTimeout(() => setIsimMsg(""), 2000); return; }
    if (benimKaydim) {
      await updateDoc(doc(db, "kullanicilar", benimKaydim.id), { ad: yeniAd });
    } else {
      await addDoc(collection(db, "kullanicilar"), {
        ad: yeniAd, email: kullanici.email, emailKucuk,
        tur: "sifreli", eklenmeTarihi: Date.now(),
      });
    }
    setIsimDuzenle(false);
    setIsimMsg("İsim güncellendi ✓");
    setTimeout(() => setIsimMsg(""), 2500);
  };

  const acikHammadde = hammaddeler.filter((h) => !h.tamamlandi);
  const acikHammaddeKg = acikHammadde.reduce((s, h) => s + (Number(h.miktar) || 0), 0);

  const dusukDepoStok = depoStok.filter((s) => s.miktar <= 0);

  const buAyMetalOlcum = metalTalepler.filter((t) => t.tarih && new Date(t.tarih).toISOString().slice(0, 7) === buAy);
  const buAyMetalKg = buAyMetalOlcum.reduce((s, t) => s + (Number(t.toplamKg) || 0), 0);

  const aktifFasonIs = fasonIsler.filter((j) => j.durum !== "tamamlandi");
  const gecikenHatirlatici = fasonHatirlaticilar.filter((r) => !r.tamamlandi && r.tarih && r.tarih < bugun);
  const bugunkuHatirlatici = fasonHatirlaticilar.filter((r) => !r.tamamlandi && r.tarih === bugun);

  const uyarilar = [];
  if (dusukDepoStok.length > 0) uyarilar.push({ metin: `${dusukDepoStok.length} depo stok kalemi tükenmiş / eksi durumda`, git: "depo-kart" });
  if (gecikenHatirlatici.length > 0) uyarilar.push({ metin: `${gecikenHatirlatici.length} fason hatırlatıcı gecikti`, git: "fason-hatirlaticilar" });
  if (bugunkuHatirlatici.length > 0) uyarilar.push({ metin: `${bugunkuHatirlatici.length} hatırlatıcının tarihi bugün`, git: "fason-hatirlaticilar" });

  const modulKartlari = [
    { id: "hammadde-kayit", baslik: "Hammadde", aciklama: "Sipariş / stok takibi", deger: `${acikHammadde.length}`, altYazi: "açık sipariş", icon: Boxes },
    { id: "metal-hizli", baslik: "Metal Ölçü", aciklama: "Kesit ağırlık hesabı", deger: `${buAyMetalKg.toLocaleString("tr-TR", { maximumFractionDigits: 0 })} kg`, altYazi: "bu ay", icon: Ruler },
    { id: "depo-kart", baslik: "Depo Stok", aciklama: "Envanter / giriş-çıkış", deger: `${depoStok.length}`, altYazi: "kalem", icon: Boxes },
    { id: "fason-ozet", baslik: "Fason Takip", aciklama: "Firma / iş / hareket", deger: `${aktifFasonIs.length}`, altYazi: "aktif iş", icon: Building2 },
  ];

  const saat = new Date().getHours();
  const selamlama = saat < 6 ? "İyi geceler" : saat < 12 ? "Günaydın" : saat < 18 ? "İyi günler" : "İyi akşamlar";

  return (
    <div style={{ display: "grid", gap: 24 }}>
      <div className="card" style={{ padding: 28, background: "linear-gradient(135deg, #1b333c 0%, #16232a 100%)" }}>
        <div style={{ fontSize: 12, color: "#2dd4bf", fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 6 }}>SAKLAZ · ÜRETİM ERP</div>
        {isimDuzenle ? (
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 4 }}>
            <input
              className="input" style={{ maxWidth: 220 }} placeholder="Adını yaz"
              value={isimGirdi} onChange={(e) => setIsimGirdi(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && isimKaydet()}
              autoFocus
            />
            <button onClick={isimKaydet} style={{ background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "8px 14px", fontWeight: 700, fontSize: 12.5, cursor: "pointer" }}>Kaydet</button>
            <button onClick={() => setIsimDuzenle(false)} className="btn-ghost">Vazgeç</button>
          </div>
        ) : (
          <div style={{ fontSize: 24, fontWeight: 800, letterSpacing: -0.3, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            {selamlama}, {gosterilenIsim}
            <button onClick={() => { setIsimGirdi(benimKaydim?.ad || ""); setIsimDuzenle(true); }} title="İsmini düzenle" style={{ background: "none", border: "1px solid #3d6169", color: "#8b929a", borderRadius: 6, padding: "3px 9px", fontSize: 11, cursor: "pointer", fontWeight: 500 }}>✎ İsmini Düzenle</button>
          </div>
        )}
        {isimMsg && <div style={{ fontSize: 12, color: "#2dd4bf", marginTop: 4 }}>{isimMsg}</div>}
        <div style={{ fontSize: 13, color: "#8b929a", marginTop: 6 }}>{new Date().toLocaleDateString("tr-TR", { weekday: "long", day: "numeric", month: "long", year: "numeric" })}</div>
      </div>

      {uyarilar.length > 0 && (
        <div style={{ display: "grid", gap: 10 }}>
          {uyarilar.map((u, i) => (
            <button key={i} onClick={() => git(u.git)} style={{ display: "flex", alignItems: "center", gap: 10, background: "#3a1f1f", border: "1px solid #5a2a2a", borderRadius: 10, padding: "12px 16px", color: "#e07a6b", fontSize: 13, cursor: "pointer", textAlign: "left" }}>
              <AlertTriangle size={16} style={{ flexShrink: 0 }} />
              {u.metin}
              <ChevronRight size={14} style={{ marginLeft: "auto", flexShrink: 0 }} />
            </button>
          ))}
        </div>
      )}

      <div>
        <div className="sect-label" style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", color: "#8b929a", fontWeight: 600, marginBottom: 12 }}>Modüller</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14 }}>
          {modulKartlari.map((k) => {
            const Icon = k.icon;
            return (
              <button key={k.id} onClick={() => git(k.id)} className="card" style={{ padding: 20, textAlign: "left", cursor: "pointer", display: "flex", flexDirection: "column", gap: 10 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <div style={{ width: 36, height: 36, borderRadius: 9, background: "#113330", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <Icon size={18} color="#2dd4bf" />
                  </div>
                  <ChevronRight size={16} color="#6b7178" />
                </div>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 14.5 }}>{k.baslik}</div>
                  <div style={{ fontSize: 11.5, color: "#8b929a", marginTop: 2 }}>{k.aciklama}</div>
                </div>
                <div style={{ marginTop: 4 }}>
                  <span style={{ fontSize: 22, fontWeight: 800, fontFamily: "monospace", color: "#2dd4bf" }}>{k.deger}</span>
                  <span style={{ fontSize: 11.5, color: "#6b7178", marginLeft: 6 }}>{k.altYazi}</span>
                </div>
              </button>
            );
          })}
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Toplam Takım" value={teams.length} />
        <Stat label="Toplam Makine" value={machines.length} />
        <Stat label="Açık Hammadde (Kg)" value={acikHammaddeKg.toLocaleString("tr-TR", { maximumFractionDigits: 0 })} />
        <Stat label="Fason Firma" value={fasonFirmalar.length} />
        <Stat label="Depo Eksi Stok" value={dusukDepoStok.length} highlight={dusukDepoStok.length > 0} />
      </div>
    </div>
  );
}

// ---------- Kayıt Ekle ----------
function KayitEkle({ teams, machines, records }) {
  const [form, setForm] = useState({ tarih: todayISO(), takim: "", magaza: "", makine: "", urun: "", adet: "" });
  const [msg, setMsg] = useState("");
  const [arama, setArama] = useState("");
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  useEffect(() => {
    if (!iceAktariliyor) return;
    const uyar = (e) => { e.preventDefault(); e.returnValue = ""; };
    window.addEventListener("beforeunload", uyar);
    return () => window.removeEventListener("beforeunload", uyar);
  }, [iceAktariliyor]);
  const [iceMsg, setIceMsg] = useState("");
  const dosyaRef = useRef(null);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const submit = async () => {
    if (!form.takim || !form.makine || !form.adet) {
      setMsg("Takım, makine ve adet zorunlu.");
      setTimeout(() => setMsg(""), 2500);
      return;
    }
    await addDoc(collection(db, "records"), { ...form, adet: Number(form.adet), olusturma: Date.now() });
    setForm((f) => ({ tarih: f.tarih, takim: "", magaza: "", makine: "", urun: "", adet: "" }));
    setMsg("Kayıt eklendi.");
    setTimeout(() => setMsg(""), 1800);
  };

  const sil = async (id) => { await deleteDoc(doc(db, "records", id)); };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true);
    setIceMsg("");
    try {
      const kayitlar = await excelDenKayitOku(dosya);
      if (kayitlar.length === 0) {
        setIceMsg("Dosyada geçerli satır bulunamadı. Takım veya Makine sütunu boş olan satırlar atlanır.");
      } else {
        const veriler = kayitlar.map((k) => ({ ...k, olusturma: Date.now() }));
        const { basarili, basarisiz } = await guvenliTopluYaz("records", veriler, (yapilan, toplam, hatali) => {
          setIceMsg(`${yapilan} / ${toplam} kayıt işlendi${hatali > 0 ? ` (${hatali} tanesi tekrar deneniyor)` : ""}…`);
        });
        if (basarisiz > 0) {
          setIceMsg(`${basarili} kayıt eklendi, ${basarisiz} kayıt eklenemedi (bağlantı sorunu). Kalanları tekrar yükleyebilirsin.`);
        } else {
          setIceMsg(`${basarili} stok kaydı içe aktarıldı.`);
        }
      }
    } catch (err) {
      console.error(err);
      setIceMsg("İçe aktarma sırasında hata oluştu: " + (err?.message || "bilinmeyen hata"));
    }
    setIceAktariliyor(false);
    e.target.value = "";
    setTimeout(() => setIceMsg(""), 7000);
  };

  const disaAktar = () => {
    excelIndir(
      records.map((r) => ({ Tarih: r.tarih, Takım: r.takim, Mağaza: r.magaza || "", Makine: r.makine, Ürün: r.urun || "", Adet: r.adet })),
      "stok-kayitlari.xlsx", "Stok Kayıtları"
    );
  };

  const listelenecek = useMemo(() => {
    const sirali = [...records].sort((a, b) => (b.olusturma || 0) - (a.olusturma || 0));
    if (!arama.trim()) return sirali.slice(0, 10);
    const q = arama.trim().toLowerCase();
    return sirali.filter((r) =>
      (r.takim || "").toLowerCase().includes(q) ||
      (r.magaza || "").toLowerCase().includes(q) ||
      (r.makine || "").toLowerCase().includes(q) ||
      (r.urun || "").toLowerCase().includes(q) ||
      (r.tarih || "").includes(q)
    ).slice(0, 50);
  }, [records, arama]);

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Günlük Çıkış Kaydı</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button
              className="btn-ghost"
              onClick={() => sablonIndir(
                ["TARİH", "TAKIM", "MAĞAZA", "MAKİNE", "ÜRÜN", "ADET"],
                [["2026-08-11", "Kesim Takım 1", "Örnek Mağaza", "Örnek Makine", "Ürün Adı", "10"]],
                "stok-kayit-sablonu.xlsx", "Şablon"
              )}
            >
              <FileDown size={14} /> Excel Şablonu İndir
            </button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}>
              <Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excel'den İçe Aktar"}
            </button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excel'e Aktar</button>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div><label className="field-label">Tarih</label><input className="input" type="date" value={form.tarih} onChange={set("tarih")} /></div>
          <div>
            <label className="field-label">Takım</label>
            <select className="input" value={form.takim} onChange={set("takim")}>
              <option value="">Seçiniz</option>
              {teams.map((t) => <option key={t.id} value={t.name}>{t.name}</option>)}
            </select>
          </div>
          <div><label className="field-label">Mağaza</label><input className="input" placeholder="Mağaza / müşteri adı" value={form.magaza} onChange={set("magaza")} /></div>
          <div>
            <label className="field-label">Makine</label>
            <select className="input" value={form.makine} onChange={set("makine")}>
              <option value="">Seçiniz</option>
              {machines.map((m) => <option key={m.id} value={m.name}>{m.name}</option>)}
            </select>
          </div>
          <div><label className="field-label">Ürün / Model</label><input className="input" placeholder="Opsiyonel" value={form.urun} onChange={set("urun")} /></div>
          <div><label className="field-label">Adet</label><input className="input" type="number" min="0" placeholder="0" value={form.adet} onChange={set("adet")} /></div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button onClick={submit} style={{ display: "flex", alignItems: "center", gap: 7, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
            <Plus size={16} /> Kaydı Ekle
          </button>
          {msg && <span style={{ fontSize: 12.5, color: "#8b929a" }}>{msg}</span>}
        </div>
        {iceMsg && (
          <div style={{ marginTop: 14, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{iceMsg}</div>
        )}
        <div style={{ fontSize: 12, color: "#6b7178", marginTop: 10 }}>
          Excel'den içe aktarırken sütun sırası: Tarih, Takım, Mağaza, Makine, Ürün, Adet. Başlık satırı olabilir.
        </div>
        {(teams.length === 0 || machines.length === 0) && (
          <div style={{ marginTop: 14, fontSize: 12.5, color: "#c98a2e", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>
            {teams.length === 0 && machines.length === 0 ? "Önce Takımlar ve Makineler sekmelerinden liste oluşturun."
              : teams.length === 0 ? "Önce Takımlar sekmesinden takım ekleyin."
              : "Önce Makineler sekmesinden makine listesini ekleyin."}
          </div>
        )}
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>{arama ? "Arama Sonuçları" : "Son Kayıtlar"}</div>
          <div style={{ position: "relative", minWidth: 220 }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input
              className="input"
              style={{ paddingLeft: 30 }}
              placeholder="Takım, mağaza, makine, tarih ara…"
              value={arama}
              onChange={(e) => setArama(e.target.value)}
            />
          </div>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr><th>Tarih</th><th>Takım</th><th>Mağaza</th><th>Makine</th><th>Ürün</th><th>Adet</th><th></th></tr></thead>
            <tbody>
              {listelenecek.length === 0 && <tr><td colSpan={7} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>{arama ? "Sonuç bulunamadı." : "Henüz kayıt yok."}</td></tr>}
              {listelenecek.map((r) => (
                <tr key={r.id}>
                  <td style={{ fontFamily: "monospace" }}>{r.tarih}</td>
                  <td>{r.takim}</td>
                  <td>{r.magaza || "—"}</td>
                  <td>{r.makine}</td>
                  <td>{r.urun || "—"}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#2dd4bf" }}>{r.adet}</td>
                  <td><button onClick={() => sil(r.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Hammadde Raporları (Stok Durumu / Sipariş Durumu) ----------
function HammaddeRaporlari({ hammaddeler }) {
  const [altTab, setAltTab] = useState("stok");
  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div style={{ display: "flex", gap: 8 }}>
        <button
          onClick={() => setAltTab("stok")}
          style={{
            flex: 1, padding: "10px 16px", borderRadius: 9, cursor: "pointer", fontWeight: 700, fontSize: 13,
            background: altTab === "stok" ? "#2dd4bf" : "#1b333c",
            color: altTab === "stok" ? "#142a30" : "#c7cbd1",
            border: `1px solid ${altTab === "stok" ? "#2dd4bf" : "#2a4b52"}`,
          }}
        >
          Hammadde Stok Raporu
        </button>
        <button
          onClick={() => setAltTab("siparis")}
          style={{
            flex: 1, padding: "10px 16px", borderRadius: 9, cursor: "pointer", fontWeight: 700, fontSize: 13,
            background: altTab === "siparis" ? "#2dd4bf" : "#1b333c",
            color: altTab === "siparis" ? "#142a30" : "#c7cbd1",
            border: `1px solid ${altTab === "siparis" ? "#2dd4bf" : "#2a4b52"}`,
          }}
        >
          Sipariş Durum Raporu
        </button>
      </div>
      {altTab === "stok" && <StokRaporu hammaddeler={hammaddeler} />}
      {altTab === "siparis" && <SiparisRaporu hammaddeler={hammaddeler} />}
    </div>
  );
}

// ================= YENİ MODÜL RAPORLARI =================

// ---------- Metal Ölçü Raporu ----------
function MetalOlcuRaporu({ metalTalepler, metalMalzemeler }) {
  const [f, setF] = useState({ arama: "", malzeme: "", tur: "", baslangic: "", bitis: "" });
  const [genisletilen, setGenisletilen] = useState(new Set());
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const filtrelenmis = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return metalTalepler.filter((t) => {
      if (f.malzeme && t.malzemeAdi !== f.malzeme) return false;
      if (f.tur && t.tur !== f.tur) return false;
      if (f.baslangic && t.tarih && new Date(t.tarih).toISOString().slice(0, 10) < f.baslangic) return false;
      if (f.bitis && t.tarih && new Date(t.tarih).toISOString().slice(0, 10) > f.bitis) return false;
      if (q && !((t.talepNo || "").toLowerCase().includes(q) || (t.malzemeAdi || "").toLowerCase().includes(q) || (t.dimLabel || "").toLowerCase().includes(q))) return false;
      return true;
    });
  }, [metalTalepler, f]);

  const malzemeGruplari = useMemo(() => {
    const map = new Map();
    filtrelenmis.forEach((t) => {
      const key = t.malzemeAdi || "Belirtilmemiş";
      if (!map.has(key)) map.set(key, { malzeme: key, kg: 0, tutar: 0, sayi: 0 });
      const g = map.get(key);
      g.kg += Number(t.toplamKg) || 0; g.tutar += Number(t.tutar) || 0; g.sayi += 1;
    });
    return [...map.values()].sort((a, b) => b.kg - a.kg);
  }, [filtrelenmis]);

  const talepGruplari = useMemo(() => {
    const map = new Map();
    filtrelenmis.forEach((t) => {
      const key = t.talepNo || "—";
      if (!map.has(key)) map.set(key, { talepNo: key, kg: 0, tutar: 0, kalem: [] });
      const g = map.get(key);
      g.kg += Number(t.toplamKg) || 0; g.tutar += Number(t.tutar) || 0; g.kalem.push(t);
    });
    return [...map.values()].sort((a, b) => b.kg - a.kg);
  }, [filtrelenmis]);

  const toplamKg = filtrelenmis.reduce((s, t) => s + (Number(t.toplamKg) || 0), 0);
  const toplamTutar = filtrelenmis.reduce((s, t) => s + (Number(t.tutar) || 0), 0);
  const grupToggle = (key) => setGenisletilen((s) => { const y = new Set(s); if (y.has(key)) y.delete(key); else y.add(key); return y; });

  const disaAktarMalzeme = () => excelIndir(malzemeGruplari.map((g) => ({ "Malzeme": g.malzeme, "Toplam Kg": g.kg.toFixed(3), "Toplam Tutar": g.tutar.toFixed(2), "Kayıt Sayısı": g.sayi })), "metal-malzeme-raporu.xlsx", "Malzeme Raporu");
  const disaAktarTalep = () => excelIndir(talepGruplari.map((g) => ({ "Talep No": g.talepNo, "Toplam Kg": g.kg.toFixed(3), "Toplam Tutar": g.tutar.toFixed(2), "Kalem Sayısı": g.kalem.length })), "metal-talep-raporu.xlsx", "Talep Raporu");

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}>Filtrele</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1", position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Talep no, malzeme, ölçü ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Malzeme</label>
            <select className="input" value={f.malzeme} onChange={setF2("malzeme")}>
              <option value="">Tümü</option>
              {metalMalzemeler.map((m) => <option key={m.id} value={m.ad}>{m.ad}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Kesit Türü</label>
            <select className="input" value={f.tur} onChange={setF2("tur")}>
              <option value="">Tümü</option>
              {KESIT_TIPLERI.map((k) => <option key={k.id} value={k.id}>{k.label}</option>)}
            </select>
          </div>
          <div><label className="field-label">Başlangıç</label><input className="input" type="date" value={f.baslangic} onChange={setF2("baslangic")} /></div>
          <div><label className="field-label">Bitiş</label><input className="input" type="date" value={f.bitis} onChange={setF2("bitis")} /></div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Toplam Kayıt" value={filtrelenmis.length} />
        <Stat label="Toplam Ağırlık" value={`${toplamKg.toLocaleString("tr-TR", { maximumFractionDigits: 2 })} kg`} highlight />
        <Stat label="Toplam Tutar" value={paraTR(toplamTutar)} />
        <Stat label="Malzeme Çeşidi" value={malzemeGruplari.length} />
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>Malzeme Bazında Dağılım ({malzemeGruplari.length})</div>
          <button className="btn-ghost" onClick={disaAktarMalzeme}><Download size={14} /> Excele Aktar</button>
        </div>
        <table>
          <thead><tr><th>Malzeme</th><th>Toplam Kg</th><th>Toplam Tutar</th><th>Kayıt Sayısı</th></tr></thead>
          <tbody>
            {malzemeGruplari.length === 0 && <tr><td colSpan={4} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Kayıt bulunamadı.</td></tr>}
            {malzemeGruplari.map((g) => (
              <tr key={g.malzeme}>
                <td>{g.malzeme}</td>
                <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#2dd4bf" }}>{g.kg.toLocaleString("tr-TR", { maximumFractionDigits: 2 })} kg</td>
                <td style={{ fontFamily: "monospace" }}>{paraTR(g.tutar)}</td>
                <td style={{ fontFamily: "monospace" }}>{g.sayi}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>Talep Bazında Dağılım ({talepGruplari.length})</div>
          <button className="btn-ghost" onClick={disaAktarTalep}><Download size={14} /> Excele Aktar</button>
        </div>
        <div style={{ overflowX: "auto", maxHeight: 480, overflowY: "auto" }}>
          <table>
            <thead><tr><th></th><th>Talep No</th><th>Toplam Kg</th><th>Toplam Tutar</th><th>Kalem</th></tr></thead>
            <tbody>
              {talepGruplari.length === 0 && <tr><td colSpan={5} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Kayıt bulunamadı.</td></tr>}
              {talepGruplari.map((g) => {
                const acik = genisletilen.has(g.talepNo);
                return (
                  <React.Fragment key={g.talepNo}>
                    <tr onClick={() => grupToggle(g.talepNo)} style={{ cursor: "pointer" }}>
                      <td>{acik ? <ChevronDown size={14} /> : <ChevronRight size={14} />}</td>
                      <td style={{ fontWeight: 700 }}>{g.talepNo}</td>
                      <td style={{ fontFamily: "monospace", color: "#2dd4bf" }}>{g.kg.toLocaleString("tr-TR", { maximumFractionDigits: 2 })} kg</td>
                      <td style={{ fontFamily: "monospace" }}>{paraTR(g.tutar)}</td>
                      <td style={{ fontFamily: "monospace" }}>{g.kalem.length}</td>
                    </tr>
                    {acik && g.kalem.map((k) => (
                      <tr key={k.id}>
                        <td></td>
                        <td style={{ paddingLeft: 20, fontSize: 12.5 }}>{k.malzemeAdi} · {KESIT_ETIKET[k.tur]} · {k.dimLabel}</td>
                        <td style={{ fontFamily: "monospace", fontSize: 12.5 }}>{k.toplamKg?.toFixed(3)} kg</td>
                        <td style={{ fontFamily: "monospace", fontSize: 12.5 }}>{k.tutar ? paraTR(k.tutar) : "—"}</td>
                        <td></td>
                      </tr>
                    ))}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Depo Stok Raporu ----------
function DepoStokRaporu({ depoStok, depoHareketler }) {
  const [f, setF] = useState({ arama: "", baslangic: "", bitis: "" });
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const donemSec = (tur) => {
    const bugun = new Date();
    if (tur === "gun") {
      const g = todayISO();
      setF((s) => ({ ...s, baslangic: g, bitis: g }));
    } else if (tur === "ay") {
      const ilkGun = new Date(bugun.getFullYear(), bugun.getMonth(), 1).toISOString().slice(0, 10);
      setF((s) => ({ ...s, baslangic: ilkGun, bitis: todayISO() }));
    } else if (tur === "yil") {
      const ilkGun = new Date(bugun.getFullYear(), 0, 1).toISOString().slice(0, 10);
      setF((s) => ({ ...s, baslangic: ilkGun, bitis: todayISO() }));
    } else {
      setF((s) => ({ ...s, baslangic: "", bitis: "" }));
    }
  };

  const filtrelenmisStok = useMemo(() => {
    if (!f.arama.trim()) return depoStok;
    const q = f.arama.trim().toLowerCase();
    return depoStok.filter((s) => s.stokKodu.toLowerCase().includes(q) || s.stokAdi.toLowerCase().includes(q));
  }, [depoStok, f.arama]);

  const filtrelenmisHareket = useMemo(() => {
    return depoHareketler.filter((h) => {
      if (f.baslangic && h.tarih && new Date(h.tarih).toISOString().slice(0, 10) < f.baslangic) return false;
      if (f.bitis && h.tarih && new Date(h.tarih).toISOString().slice(0, 10) > f.bitis) return false;
      return true;
    });
  }, [depoHareketler, f]);

  const toplamGiris = filtrelenmisHareket.filter((h) => h.tip === "giris").reduce((s, h) => s + (h.miktar || 0), 0);
  const toplamCikis = filtrelenmisHareket.filter((h) => h.tip === "cikis").reduce((s, h) => s + (h.miktar || 0), 0);
  const dusukStok = depoStok.filter((s) => s.miktar <= 0).length;
  const enCokHareketGorenler = useMemo(() => {
    const map = new Map();
    filtrelenmisHareket.forEach((h) => {
      const key = h.stokKodu;
      if (!map.has(key)) map.set(key, { stokKodu: h.stokKodu, stokAdi: h.stokAdi, giris: 0, cikis: 0, hareketSayisi: 0 });
      const g = map.get(key);
      if (h.tip === "giris") g.giris += h.miktar || 0; else g.cikis += h.miktar || 0;
      g.hareketSayisi += 1;
    });
    return [...map.values()].sort((a, b) => (b.giris + b.cikis) - (a.giris + a.cikis));
  }, [filtrelenmisHareket]);

  const makineBazliCikis = useMemo(() => {
    const map = new Map();
    filtrelenmisHareket.filter((h) => h.tip === "cikis" && h.hedefMakine).forEach((h) => {
      const key = h.hedefMakine;
      if (!map.has(key)) map.set(key, { makine: key, toplam: 0, hareketSayisi: 0 });
      const g = map.get(key);
      g.toplam += h.miktar || 0; g.hareketSayisi += 1;
    });
    return [...map.values()].sort((a, b) => b.toplam - a.toplam);
  }, [filtrelenmisHareket]);

  const disaAktarStok = () => excelIndir(filtrelenmisStok.map((s) => ({ "Stok Kodu": s.stokKodu, "Stok Adı": s.stokAdi, "Miktar": s.miktar, "Birim": s.birim || "Adet" })), "depo-stok-raporu.xlsx", "Depo Stok");
  const disaAktarHareket = () => excelIndir(enCokHareketGorenler.map((g) => ({ "Stok Kodu": g.stokKodu, "Stok Adı": g.stokAdi, "Toplam Giriş": g.giris, "Toplam Çıkış": g.cikis, "Hareket Sayısı": g.hareketSayisi })), "depo-hareket-ozet-raporu.xlsx", "Hareket Özeti");
  const disaAktarMakine = () => excelIndir(makineBazliCikis.map((g) => ({ "Makine": g.makine, "Toplam Çıkış": g.toplam, "Hareket Sayısı": g.hareketSayisi })), "depo-makine-cikis-raporu.xlsx", "Makine Bazlı Çıkış");

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Filtrele</div>
          <button className="btn-ghost" onClick={disaAktarStok}><Download size={14} /> Excele Aktar</button>
        </div>
        <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
          <button className="btn-ghost" onClick={() => donemSec("gun")}>Bugün</button>
          <button className="btn-ghost" onClick={() => donemSec("ay")}>Bu Ay</button>
          <button className="btn-ghost" onClick={() => donemSec("yil")}>Bu Yıl</button>
          <button className="btn-ghost" onClick={() => donemSec("tumu")}>Tüm Zamanlar</button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Stok kodu / adı ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div><label className="field-label">Hareket Başlangıç</label><input className="input" type="date" value={f.baslangic} onChange={setF2("baslangic")} /></div>
          <div><label className="field-label">Hareket Bitiş</label><input className="input" type="date" value={f.bitis} onChange={setF2("bitis")} /></div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Toplam Kalem" value={depoStok.length} />
        <Stat label="Stokta Biten / Eksi" value={dusukStok} highlight={dusukStok > 0} />
        <Stat label="Dönem Giriş" value={toplamGiris.toLocaleString("tr-TR")} highlight />
        <Stat label="Dönem Çıkış" value={toplamCikis.toLocaleString("tr-TR")} />
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>Makine Bazında Çıkış ({makineBazliCikis.length})</div>
          <button className="btn-ghost" onClick={disaAktarMakine}><Download size={14} /> Excele Aktar</button>
        </div>
        <div style={{ overflowX: "auto", maxHeight: 400, overflowY: "auto" }}>
          <table>
            <thead><tr><th>Makine</th><th>Toplam Çıkış</th><th>Hareket Sayısı</th></tr></thead>
            <tbody>
              {makineBazliCikis.length === 0 && <tr><td colSpan={3} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Bu dönemde makineye çıkış yok.</td></tr>}
              {makineBazliCikis.map((g) => (
                <tr key={g.makine}>
                  <td><span className="pill">{g.makine}</span></td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#e07a6b" }}>{g.toplam.toLocaleString("tr-TR")}</td>
                  <td style={{ fontFamily: "monospace" }}>{g.hareketSayisi}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>Güncel Stok Durumu ({filtrelenmisStok.length})</div>
          <button className="btn-ghost" onClick={disaAktarStok}><Download size={14} /> Excele Aktar</button>
        </div>
        <div style={{ overflowX: "auto", maxHeight: 400, overflowY: "auto" }}>
          <table>
            <thead><tr><th>Stok Kodu</th><th>Stok Adı</th><th>Miktar</th><th>Birim</th></tr></thead>
            <tbody>
              {filtrelenmisStok.length === 0 && <tr><td colSpan={4} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Kayıt bulunamadı.</td></tr>}
              {filtrelenmisStok.map((s) => (
                <tr key={s.id}>
                  <td style={{ fontFamily: "monospace" }}>{s.stokKodu}</td>
                  <td>{s.stokAdi}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: s.miktar <= 0 ? "#e07a6b" : "#2dd4bf" }}>{s.miktar}</td>
                  <td>{s.birim || "Adet"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>Kalem Bazında Hareket Özeti ({enCokHareketGorenler.length})</div>
          <button className="btn-ghost" onClick={disaAktarHareket}><Download size={14} /> Excele Aktar</button>
        </div>
        <div style={{ overflowX: "auto", maxHeight: 400, overflowY: "auto" }}>
          <table>
            <thead><tr><th>Stok Kodu</th><th>Stok Adı</th><th>Toplam Giriş</th><th>Toplam Çıkış</th><th>Hareket Sayısı</th></tr></thead>
            <tbody>
              {enCokHareketGorenler.length === 0 && <tr><td colSpan={5} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Bu dönemde hareket yok.</td></tr>}
              {enCokHareketGorenler.map((g) => (
                <tr key={g.stokKodu}>
                  <td style={{ fontFamily: "monospace" }}>{g.stokKodu}</td>
                  <td>{g.stokAdi}</td>
                  <td style={{ fontFamily: "monospace", color: "#2dd4bf" }}>{g.giris}</td>
                  <td style={{ fontFamily: "monospace", color: "#e07a6b" }}>{g.cikis}</td>
                  <td style={{ fontFamily: "monospace" }}>{g.hareketSayisi}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Fason Takip Raporu ----------
function FasonTakipRaporu({ fasonFirmalar, fasonIsler, fasonHareketler }) {
  const [f, setF] = useState({ arama: "", firmaId: "" });
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const filtrelenmisIsler = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return fasonIsler.filter((j) => {
      if (f.firmaId && j.firmaId !== f.firmaId) return false;
      if (q && !((j.projeKodu || "").toLowerCase().includes(q) || (j.projeAdi || "").toLowerCase().includes(q))) return false;
      return true;
    });
  }, [fasonIsler, f]);

  const isIdSet = new Set(filtrelenmisIsler.map((j) => j.id));
  const ilgiliHareketler = fasonHareketler.filter((m) => isIdSet.has(m.isId));

  const toplamGiden = ilgiliHareketler.filter((m) => m.tip === "giden").reduce((s, m) => s + (Number(m.miktar) || 0) * (Number(m.birimFiyat) || 0), 0);
  const toplamGelen = ilgiliHareketler.filter((m) => m.tip === "gelen").reduce((s, m) => s + (Number(m.miktar) || 0) * (Number(m.birimFiyat) || 0), 0);

  const durumDagilimi = useMemo(() => {
    const c = { bekliyor: 0, uretimde: 0, tamamlandi: 0 };
    filtrelenmisIsler.forEach((j) => { c[j.durum || "bekliyor"] = (c[j.durum || "bekliyor"] || 0) + 1; });
    return c;
  }, [filtrelenmisIsler]);

  const kaliteDagilimi = useMemo(() => {
    const c = { okeylendi: 0, red: 0, olcumde: 0, belirsiz: 0 };
    filtrelenmisIsler.forEach((j) => { c[j.kaliteDurumu || "belirsiz"] = (c[j.kaliteDurumu || "belirsiz"] || 0) + 1; });
    return c;
  }, [filtrelenmisIsler]);

  const hammaddeGonderildiMi = (isId) => fasonHareketler.some((m) => m.isId === isId && m.tip === "giden");
  const gonderilenSayisi = filtrelenmisIsler.filter((j) => hammaddeGonderildiMi(j.id)).length;

  const firmaDetay = useMemo(() => {
    return fasonFirmalar.filter((firma) => !f.firmaId || firma.id === f.firmaId).map((firma) => {
      const isler = filtrelenmisIsler.filter((j) => j.firmaId === firma.id);
      const isIdler = new Set(isler.map((j) => j.id));
      let giden = 0, gelen = 0;
      fasonHareketler.forEach((m) => {
        if (!isIdler.has(m.isId)) return;
        const t = (Number(m.miktar) || 0) * (Number(m.birimFiyat) || 0);
        if (m.tip === "giden") giden += t; else gelen += t;
      });
      return { firma, isSayisi: isler.length, aktifIsSayisi: isler.filter((j) => j.durum !== "tamamlandi").length, giden, gelen, bakiye: giden - gelen };
    }).filter((r) => r.isSayisi > 0 || !f.firmaId).sort((a, b) => b.bakiye - a.bakiye);
  }, [fasonFirmalar, filtrelenmisIsler, fasonHareketler, f.firmaId]);

  const disaAktarFirma = () => excelIndir(firmaDetay.map((r) => ({ "Firma": r.firma.ad, "Toplam İş": r.isSayisi, "Aktif İş": r.aktifIsSayisi, "Giden": r.giden.toFixed(2), "Gelen": r.gelen.toFixed(2), "Bakiye": r.bakiye.toFixed(2) })), "fason-firma-raporu.xlsx", "Firma Raporu");

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}>Filtrele</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Proje kodu / adı ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Firma</label>
            <select className="input" value={f.firmaId} onChange={setF2("firmaId")}>
              <option value="">Tümü</option>
              {fasonFirmalar.map((fm) => <option key={fm.id} value={fm.id}>{fm.ad}{fm.not && fm.not.trim() ? ` — ${fm.not}` : ""}</option>)}
            </select>
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Toplam İş" value={filtrelenmisIsler.length} />
        <Stat label="Toplam Giden" value={paraTR(toplamGiden)} />
        <Stat label="Toplam Gelen" value={paraTR(toplamGelen)} />
        <Stat label="Net Bakiye" value={paraTR(toplamGiden - toplamGelen)} highlight />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14 }}>
        <div className="card" style={{ padding: 20 }}>
          <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 12 }}>İş Durum Dağılımı</div>
          {Object.entries(FASON_DURUM).map(([k, d]) => (
            <div key={k} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "6px 0", borderBottom: "1px solid #223b42" }}>
              <span style={{ fontSize: 13, display: "flex", alignItems: "center", gap: 8 }}><span style={{ width: 10, height: 10, borderRadius: 3, background: d.renk, display: "inline-block" }} />{d.label}</span>
              <span style={{ fontFamily: "monospace", fontWeight: 700 }}>{durumDagilimi[k] || 0}</span>
            </div>
          ))}
        </div>
        <div className="card" style={{ padding: 20 }}>
          <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 12 }}>Kalite Kontrolü Dağılımı</div>
          {Object.entries(FASON_KALITE).map(([k, q]) => (
            <div key={k} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "6px 0", borderBottom: "1px solid #223b42" }}>
              <span style={{ fontSize: 13, display: "flex", alignItems: "center", gap: 8 }}><span style={{ width: 10, height: 10, borderRadius: 3, background: q.renk, display: "inline-block" }} />{q.label}</span>
              <span style={{ fontFamily: "monospace", fontWeight: 700 }}>{kaliteDagilimi[k] || 0}</span>
            </div>
          ))}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "6px 0" }}>
            <span style={{ fontSize: 13, color: "#8b929a" }}>Belirsiz</span>
            <span style={{ fontFamily: "monospace", fontWeight: 700, color: "#8b929a" }}>{kaliteDagilimi.belirsiz || 0}</span>
          </div>
        </div>
        <div className="card" style={{ padding: 20 }}>
          <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 12 }}>Hammadde Gönderim Oranı</div>
          <div style={{ fontSize: 26, fontWeight: 800, fontFamily: "monospace", color: "#2dd4bf" }}>{gonderilenSayisi} / {filtrelenmisIsler.length}</div>
          <div style={{ fontSize: 12, color: "#8b929a", marginTop: 6 }}>işe hammadde gönderilmiş</div>
        </div>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>Firma Bazında Detay ({firmaDetay.length})</div>
          <button className="btn-ghost" onClick={disaAktarFirma}><Download size={14} /> Excele Aktar</button>
        </div>
        <div style={{ overflowX: "auto", maxHeight: 480, overflowY: "auto" }}>
          <table>
            <thead><tr><th>Firma</th><th>Toplam İş</th><th>Aktif İş</th><th>Giden</th><th>Gelen</th><th>Bakiye</th></tr></thead>
            <tbody>
              {firmaDetay.length === 0 && <tr><td colSpan={6} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Kayıt bulunamadı.</td></tr>}
              {firmaDetay.map((r) => (
                <tr key={r.firma.id}>
                  <td>{r.firma.ad}</td>
                  <td style={{ fontFamily: "monospace" }}>{r.isSayisi}</td>
                  <td style={{ fontFamily: "monospace" }}>{r.aktifIsSayisi}</td>
                  <td style={{ fontFamily: "monospace", color: "#e8a33d" }}>{paraTR(r.giden)}</td>
                  <td style={{ fontFamily: "monospace", color: "#4b8f5e" }}>{paraTR(r.gelen)}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: r.bakiye >= 0 ? "#2dd4bf" : "#e07a6b" }}>{paraTR(r.bakiye)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Üretim Raporu ----------
function UretimRaporu({ teams, machines, records }) {
  const [f, setF] = useState({ start: "", end: "", takim: "", makine: "", magaza: "", arama: "" });
  const set = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const filtered = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return records.filter((r) => {
      if (f.start && r.tarih < f.start) return false;
      if (f.end && r.tarih > f.end) return false;
      if (f.takim && r.takim !== f.takim) return false;
      if (f.makine && r.makine !== f.makine) return false;
      if (f.magaza && !(r.magaza || "").toLowerCase().includes(f.magaza.toLowerCase())) return false;
      if (q && !(
        (r.takim || "").toLowerCase().includes(q) ||
        (r.magaza || "").toLowerCase().includes(q) ||
        (r.makine || "").toLowerCase().includes(q) ||
        (r.urun || "").toLowerCase().includes(q)
      )) return false;
      return true;
    }).sort((a, b) => (a.tarih < b.tarih ? 1 : -1));
  }, [records, f]);

  const sil = async (id) => { await deleteDoc(doc(db, "records", id)); };
  const toplam = filtered.reduce((s, r) => s + (Number(r.adet) || 0), 0);
  const magazalar = [...new Set(records.map((r) => r.magaza).filter(Boolean))];

  const disaAktar = () => {
    excelIndir(
      filtered.map((r) => ({ Tarih: r.tarih, Takım: r.takim, Mağaza: r.magaza || "", Makine: r.makine, Ürün: r.urun || "", Adet: r.adet })),
      "uretim-raporu.xlsx",
      "Rapor"
    );
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Filtrele</div>
          <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excel'e Aktar</button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1" }}>
            <label className="field-label">Serbest Arama</label>
            <div style={{ position: "relative" }}>
              <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
              <input className="input" style={{ paddingLeft: 30 }} placeholder="Takım, mağaza, makine, ürün ara…" value={f.arama} onChange={set("arama")} />
            </div>
          </div>
          <div><label className="field-label">Başlangıç</label><input className="input" type="date" value={f.start} onChange={set("start")} /></div>
          <div><label className="field-label">Bitiş</label><input className="input" type="date" value={f.end} onChange={set("end")} /></div>
          <div>
            <label className="field-label">Takım</label>
            <select className="input" value={f.takim} onChange={set("takim")}>
              <option value="">Tümü</option>
              {teams.map((t) => <option key={t.id} value={t.name}>{t.name}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Makine</label>
            <select className="input" value={f.makine} onChange={set("makine")}>
              <option value="">Tümü</option>
              {machines.map((m) => <option key={m.id} value={m.name}>{m.name}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Mağaza</label>
            <input className="input" list="magaza-list" placeholder="Ara / seç" value={f.magaza} onChange={set("magaza")} />
            <datalist id="magaza-list">{magazalar.map((m) => <option key={m} value={m} />)}</datalist>
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Kayıt Sayısı" value={filtered.length} />
        <Stat label="Toplam Çıkış (Adet)" value={toplam.toLocaleString("tr-TR")} highlight />
        <Stat label="Aktif Mağaza" value={new Set(filtered.map((r) => r.magaza).filter(Boolean)).size} />
        <Stat label="Aktif Makine" value={new Set(filtered.map((r) => r.makine)).size} />
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Sonuçlar ({filtered.length})</div>
        <div style={{ overflowX: "auto", maxHeight: 480, overflowY: "auto" }}>
          <table>
            <thead><tr><th>Tarih</th><th>Takım</th><th>Mağaza</th><th>Makine</th><th>Ürün</th><th>Adet</th><th></th></tr></thead>
            <tbody>
              {filtered.length === 0 && <tr><td colSpan={7} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Sonuç bulunamadı.</td></tr>}
              {filtered.map((r) => (
                <tr key={r.id}>
                  <td style={{ fontFamily: "monospace" }}>{r.tarih}</td>
                  <td>{r.takim}</td>
                  <td>{r.magaza || "—"}</td>
                  <td>{r.makine}</td>
                  <td>{r.urun || "—"}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#2dd4bf" }}>{r.adet}</td>
                  <td><button onClick={() => sil(r.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function Stat({ label, value, highlight }) {
  return (
    <div className="card" style={{ padding: "16px 18px" }}>
      <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", color: "#8b929a", fontWeight: 600, marginBottom: 8 }}>{label}</div>
      <div style={{ fontSize: 26, fontWeight: 800, fontFamily: "monospace", color: highlight ? "#2dd4bf" : "#e7e5e0" }}>{value}</div>
    </div>
  );
}

// ---------- Hammadde Stok Raporu ----------
function StokRaporu({ hammaddeler }) {
  const [f, setF] = useState({ arama: "", kalite: "", cari: "" });
  const [genisletilenler, setGenisletilenler] = useState(new Set());
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const kaliteler = [...new Set(hammaddeler.map((h) => h.kalite).filter(Boolean))];
  const cariler = [...new Set(hammaddeler.map((h) => h.cari).filter(Boolean))];

  // Stokta sayılan: henüz tamamlanmamış (kullanılmamış / tükenmemiş) kayıtlar
  const acikKayitlar = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return hammaddeler.filter((h) => {
      if (h.tamamlandi) return false;
      if (f.kalite && h.kalite !== f.kalite) return false;
      if (f.cari && h.cari !== f.cari) return false;
      if (q && !(
        (h.kalite || "").toLowerCase().includes(q) ||
        (h.cari || "").toLowerCase().includes(q) ||
        (h.aciklama1 || "").toLowerCase().includes(q) ||
        (h.aciklama2 || "").toLowerCase().includes(q) ||
        (h.projeKodu || "").toLowerCase().includes(q) ||
        (h.projeAdi || "").toLowerCase().includes(q)
      )) return false;
      return true;
    });
  }, [hammaddeler, f]);

  const kaliteGruplari = useMemo(() => {
    const map = new Map();
    for (const h of acikKayitlar) {
      const anahtar = h.kalite || "Belirtilmemiş";
      if (!map.has(anahtar)) map.set(anahtar, { kalite: anahtar, adet: 0, kayitSayisi: 0, kayitlar: [] });
      const g = map.get(anahtar);
      g.adet += Number(h.miktar) || 0;
      g.kayitSayisi += 1;
      g.kayitlar.push(h);
    }
    return [...map.values()].sort((a, b) => b.adet - a.adet);
  }, [acikKayitlar]);

  // Arama yapılırken eşleşen gruplar otomatik açılsın
  useEffect(() => {
    if (f.arama.trim()) {
      setGenisletilenler(new Set(kaliteGruplari.map((g) => g.kalite)));
    }
  }, [f.arama]);

  const grupToggle = (kalite) => {
    setGenisletilenler((s) => {
      const y = new Set(s);
      if (y.has(kalite)) y.delete(kalite); else y.add(kalite);
      return y;
    });
  };

  const toplamAdet = acikKayitlar.reduce((s, h) => s + (Number(h.miktar) || 0), 0);

  const disaAktar = () => {
    excelIndir(
      kaliteGruplari.map((g) => ({ "Kalite / Malzeme": g.kalite, "Toplam Kg": g.adet, "Kayıt Sayısı": g.kayitSayisi })),
      "hammadde-stok-raporu.xlsx", "Stok Raporu"
    );
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Filtrele</div>
          <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excel'e Aktar</button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1", position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Kalite, cari, ölçü (Ø35 gibi), parça, proje ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Kalite</label>
            <select className="input" value={f.kalite} onChange={setF2("kalite")}>
              <option value="">Tümü</option>
              {kaliteler.map((k) => <option key={k} value={k}>{k}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Cari</label>
            <select className="input" value={f.cari} onChange={setF2("cari")}>
              <option value="">Tümü</option>
              {cariler.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
        </div>
        <div style={{ fontSize: 12, color: "#6b7178", marginTop: 10 }}>
          Bu rapor, henüz "Tamamlandı" işaretlenmemiş (Açık Siparişler'deki) hammadde kayıtlarını, Kalite / malzeme türüne göre gruplayıp toplam adedini gösterir. Bir satıra tıklayınca o kaliteye ait tüm kayıtlar açılır.
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Malzeme Türü Sayısı" value={kaliteGruplari.length} />
        <Stat label="Toplam Stok (Kg)" value={toplamAdet.toLocaleString("tr-TR")} highlight />
        <Stat label="Açık Kayıt Sayısı" value={acikKayitlar.length} />
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Malzeme Bazında Stok ({kaliteGruplari.length})</div>
        <div style={{ overflowX: "auto", maxHeight: 600, overflowY: "auto" }}>
          <table>
            <thead><tr><th style={{ width: 24 }}></th><th>Kalite / Malzeme</th><th>Toplam Kg</th><th>Kayıt Sayısı</th></tr></thead>
            <tbody>
              {kaliteGruplari.length === 0 && <tr><td colSpan={4} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Sonuç bulunamadı.</td></tr>}
              {kaliteGruplari.map((g) => {
                const acik = genisletilenler.has(g.kalite);
                return (
                  <React.Fragment key={g.kalite}>
                    <tr onClick={() => grupToggle(g.kalite)} style={{ cursor: "pointer" }}>
                      <td style={{ color: "#6b7178" }}>{acik ? <ChevronDown size={14} /> : <ChevronRight size={14} />}</td>
                      <td style={{ fontWeight: 700 }}>{g.kalite}</td>
                      <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#2dd4bf" }}>{g.adet.toLocaleString("tr-TR")} kg</td>
                      <td style={{ fontFamily: "monospace" }}>{g.kayitSayisi}</td>
                    </tr>
                    {acik && (
                      <tr>
                        <td></td>
                        <td colSpan={3} style={{ padding: 0, background: "#1c222b" }}>
                          <table style={{ margin: 0 }}>
                            <thead>
                              <tr>
                                <th style={{ paddingLeft: 20 }}>Cari</th>
                                <th>Proje Kodu</th>
                                <th>Açıklama 1 (Ölçü)</th>
                                <th>Açıklama 2 (Parça)</th>
                                <th>Kg</th>
                              </tr>
                            </thead>
                            <tbody>
                              {g.kayitlar.map((h) => (
                                <tr key={h.id}>
                                  <td style={{ paddingLeft: 20 }}>{h.cari}</td>
                                  <td style={{ fontFamily: "monospace" }}>{h.projeKodu || "—"}</td>
                                  <td>{h.aciklama1 || "—"}</td>
                                  <td>{h.aciklama2 || "—"}</td>
                                  <td style={{ fontFamily: "monospace", color: "#2dd4bf" }}>{(h.miktar || 0).toLocaleString("tr-TR")}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Sipariş Durum Raporu ----------
function SiparisRaporu({ hammaddeler }) {
  const [f, setF] = useState({ arama: "", cari: "", projeKodu: "" });
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const cariler = [...new Set(hammaddeler.map((h) => h.cari).filter(Boolean))];
  const projeler = [...new Set(hammaddeler.map((h) => h.projeKodu).filter(Boolean))];

  const filtrelenmis = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return hammaddeler.filter((h) => {
      if (f.cari && h.cari !== f.cari) return false;
      if (f.projeKodu && h.projeKodu !== f.projeKodu) return false;
      if (q && !(
        (h.cari || "").toLowerCase().includes(q) ||
        (h.projeKodu || "").toLowerCase().includes(q) ||
        (h.projeAdi || "").toLowerCase().includes(q)
      )) return false;
      return true;
    });
  }, [hammaddeler, f]);

  const acikSayisi = filtrelenmis.filter((h) => !h.tamamlandi).length;
  const tamamlananSayisi = filtrelenmis.filter((h) => h.tamamlandi).length;

  const cariGruplari = useMemo(() => {
    const map = new Map();
    for (const h of filtrelenmis) {
      const anahtar = h.cari || "Belirtilmemiş";
      if (!map.has(anahtar)) map.set(anahtar, { cari: anahtar, acik: 0, tamamlanan: 0 });
      const g = map.get(anahtar);
      if (h.tamamlandi) g.tamamlanan += 1; else g.acik += 1;
    }
    return [...map.values()].sort((a, b) => (b.acik + b.tamamlanan) - (a.acik + a.tamamlanan));
  }, [filtrelenmis]);

  const disaAktar = () => {
    excelIndir(
      cariGruplari.map((g) => ({ "Cari İsmi": g.cari, "Açık Sipariş": g.acik, "Tamamlanan": g.tamamlanan, "Toplam": g.acik + g.tamamlanan })),
      "siparis-durum-raporu.xlsx", "Sipariş Durumu"
    );
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Filtrele</div>
          <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excel'e Aktar</button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1", position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Cari, proje ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Cari</label>
            <select className="input" value={f.cari} onChange={setF2("cari")}>
              <option value="">Tümü</option>
              {cariler.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Proje Kodu</label>
            <select className="input" value={f.projeKodu} onChange={setF2("projeKodu")}>
              <option value="">Tümü</option>
              {projeler.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Açık Sipariş" value={acikSayisi} />
        <Stat label="Tamamlanan" value={tamamlananSayisi} highlight />
        <Stat label="Toplam Kayıt" value={filtrelenmis.length} />
        <Stat label="Firma Sayısı" value={cariGruplari.length} />
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Cari Bazında Sipariş Durumu ({cariGruplari.length})</div>
        <div style={{ overflowX: "auto", maxHeight: 480, overflowY: "auto" }}>
          <table>
            <thead><tr><th>Cari İsmi</th><th>Açık Sipariş</th><th>Tamamlanan</th><th>Toplam</th></tr></thead>
            <tbody>
              {cariGruplari.length === 0 && <tr><td colSpan={4} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Sonuç bulunamadı.</td></tr>}
              {cariGruplari.map((g) => (
                <tr key={g.cari}>
                  <td>{g.cari}</td>
                  <td style={{ fontFamily: "monospace", color: "#2dd4bf", fontWeight: 700 }}>{g.acik}</td>
                  <td style={{ fontFamily: "monospace" }}>{g.tamamlanan}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700 }}>{g.acik + g.tamamlanan}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Hammadde Takip ----------
const DURUM_SECENEKLERI = ["Sipariş Verildi", "Yolda", "Depoda", "Kullanıldı"];

function HammaddeTakip({ hammaddeler }) {
  const [gorunum, setGorunum] = useState("acik"); // "acik" | "tamamlanan"
  const [form, setForm] = useState({ cari: "", projeKodu: "", projeAdi: "", kalite: "", aciklama1: "", aciklama2: "", miktar: "", durumu: "" });
  const [msg, setMsg] = useState("");
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  useEffect(() => {
    if (!iceAktariliyor) return;
    const uyar = (e) => { e.preventDefault(); e.returnValue = ""; };
    window.addEventListener("beforeunload", uyar);
    return () => window.removeEventListener("beforeunload", uyar);
  }, [iceAktariliyor]);
  const [iceMsg, setIceMsg] = useState("");
  const [f, setF] = useState({ arama: "", cari: "", projeKodu: "", durumu: "" });
  const [secililer, setSecililer] = useState(new Set());
  const dosyaRef = useRef(null);
  const set = (k) => (e) => setForm((s) => ({ ...s, [k]: e.target.value }));
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const submit = async () => {
    if (!form.cari || !form.aciklama2) {
      setMsg("Cari isim ve açıklama (parça) zorunlu.");
      setTimeout(() => setMsg(""), 2500);
      return;
    }
    await addDoc(collection(db, "hammadde"), { ...form, miktar: Number(form.miktar) || 0, tamamlandi: gorunum === "tamamlanan", olusturma: Date.now() });
    setForm({ cari: "", projeKodu: "", projeAdi: "", kalite: "", aciklama1: "", aciklama2: "", miktar: "", durumu: "" });
    setMsg(`Hammadde kaydı ${gorunum === "tamamlanan" ? "Tamamlanan" : "Açık Siparişler"} listesine eklendi.`);
    setTimeout(() => setMsg(""), 1800);
  };

  const sil = async (id) => { await deleteDoc(doc(db, "hammadde", id)); };
  const durumDegistir = async (id, deger) => { await updateDoc(doc(db, "hammadde", id), { tamamlandi: deger }); };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true);
    setIceMsg("");
    try {
      const kayitlar = await excelDenHammaddeOku(dosya);
      if (kayitlar.length === 0) {
        setIceMsg("Dosyada geçerli satır bulunamadı. Cari İsmi boş olan satırlar atlanır.");
      } else {
        const hedefTamamlandi = gorunum === "tamamlanan";
        const veriler = kayitlar.map((k) => ({ ...k, tamamlandi: hedefTamamlandi, olusturma: Date.now() }));
        const { basarili, basarisiz } = await guvenliTopluYaz("hammadde", veriler, (yapilan, toplam, hatali) => {
          setIceMsg(`${yapilan} / ${toplam} kayıt işlendi${hatali > 0 ? ` (${hatali} tanesi tekrar deneniyor)` : ""}…`);
        });
        if (basarisiz > 0) {
          setIceMsg(`${basarili} kayıt eklendi, ${basarisiz} kayıt eklenemedi (bağlantı sorunu). Aynı dosyayı tekrar yükleyip deneyebilirsin, zaten eklenenler tekrar eklenmez diye kontrol yapılmaz — dikkatli ol.`);
        } else {
          setIceMsg(`${basarili} hammadde kaydı içe aktarıldı (${hedefTamamlandi ? "Tamamlanan" : "Açık Siparişler"}'e eklendi).`);
        }
      }
    } catch (err) {
      console.error(err);
      setIceMsg("İçe aktarma sırasında hata oluştu: " + (err?.message || "bilinmeyen hata"));
    }
    setIceAktariliyor(false);
    e.target.value = "";
    setTimeout(() => setIceMsg(""), 9000);
  };

  const disaAktar = () => {
    excelIndir(
      filtrelenmis.map((h) => ({
        "CARİ İSMİ": h.cari, "PROJE KODU": h.projeKodu, "PROJE ADI": h.projeAdi,
        "KALİTE": h.kalite, "AÇIKLAMA 1": h.aciklama1, "AÇIKLAMA 2": h.aciklama2, "MİKTAR (KG)": h.miktar || 0, "DURUMU": h.durumu,
      })),
      gorunum === "acik" ? "hammadde-acik-siparisler.xlsx" : "hammadde-tamamlanan.xlsx",
      "Hammadde"
    );
  };

  const cariler = [...new Set(hammaddeler.map((h) => h.cari).filter(Boolean))];
  const projeler = [...new Set(hammaddeler.map((h) => h.projeKodu).filter(Boolean))];
  const acikSayisi = hammaddeler.filter((h) => !h.tamamlandi).length;
  const tamamlananSayisi = hammaddeler.filter((h) => h.tamamlandi).length;

  const filtrelenmis = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return hammaddeler.filter((h) => {
      if (gorunum === "acik" && h.tamamlandi) return false;
      if (gorunum === "tamamlanan" && !h.tamamlandi) return false;
      if (f.cari && h.cari !== f.cari) return false;
      if (f.projeKodu && h.projeKodu !== f.projeKodu) return false;
      if (f.durumu && h.durumu !== f.durumu) return false;
      if (q && !(
        (h.cari || "").toLowerCase().includes(q) ||
        (h.projeKodu || "").toLowerCase().includes(q) ||
        (h.projeAdi || "").toLowerCase().includes(q) ||
        (h.kalite || "").toLowerCase().includes(q) ||
        (h.aciklama1 || "").toLowerCase().includes(q) ||
        (h.aciklama2 || "").toLowerCase().includes(q)
      )) return false;
      return true;
    }).sort((a, b) => (b.olusturma || 0) - (a.olusturma || 0));
  }, [hammaddeler, f, gorunum]);

  // Görünüm değişince seçim listesini temizle
  useEffect(() => { setSecililer(new Set()); }, [gorunum]);

  const hepsiSecili = filtrelenmis.length > 0 && filtrelenmis.every((h) => secililer.has(h.id));
  const tumunuSecToggle = () => {
    if (hepsiSecili) setSecililer(new Set());
    else setSecililer(new Set(filtrelenmis.map((h) => h.id)));
  };
  const birSecToggle = (id) => {
    setSecililer((s) => {
      const yeni = new Set(s);
      if (yeni.has(id)) yeni.delete(id); else yeni.add(id);
      return yeni;
    });
  };

  const topluIsaretle = async (deger) => {
    for (const id of secililer) {
      await updateDoc(doc(db, "hammadde", id), { tamamlandi: deger });
    }
    setSecililer(new Set());
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      {/* Açık / Tamamlanan görünüm anahtarı - formun hangi listeye kayıt atacağını belirler */}
      <div style={{ display: "flex", gap: 8 }}>
        <button
          onClick={() => setGorunum("acik")}
          style={{
            flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
            padding: "12px 16px", borderRadius: 9, cursor: "pointer", fontWeight: 700, fontSize: 13.5,
            background: gorunum === "acik" ? "#2dd4bf" : "#1b333c",
            color: gorunum === "acik" ? "#142a30" : "#c7cbd1",
            border: `1px solid ${gorunum === "acik" ? "#2dd4bf" : "#2a4b52"}`,
          }}
        >
          Açık Siparişler <span style={{ opacity: 0.75 }}>({acikSayisi})</span>
        </button>
        <button
          onClick={() => setGorunum("tamamlanan")}
          style={{
            flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
            padding: "12px 16px", borderRadius: 9, cursor: "pointer", fontWeight: 700, fontSize: 13.5,
            background: gorunum === "tamamlanan" ? "#2dd4bf" : "#1b333c",
            color: gorunum === "tamamlanan" ? "#142a30" : "#c7cbd1",
            border: `1px solid ${gorunum === "tamamlanan" ? "#2dd4bf" : "#2a4b52"}`,
          }}
        >
          Tamamlanan <span style={{ opacity: 0.75 }}>({tamamlananSayisi})</span>
        </button>
      </div>

      <div className="card" style={{ padding: 20, borderColor: gorunum === "tamamlanan" ? "#3a3220" : undefined }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>
            {gorunum === "tamamlanan" ? "Tamamlanan Sipariş Kaydı Ekle" : "Açık Sipariş Kaydı Ekle"}
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button
              className="btn-ghost"
              onClick={() => sablonIndir(
                ["CARİ İSMİ", "PROJE KODU", "PROJE ADI", "KALİTE", "AÇIKLAMA 1", "AÇIKLAMA 2", "MİKTAR (KG)", "DURUMU"],
                [["ÖRNEK FİRMA A.Ş.", "2026-092", "ENDERUS", "BRONZ METAL", "Ø30X375 1 ADET 10 KALAY", "QSB NAMLU KOUMLAMA AP. - 1", "10", ""]],
                "hammadde-sablonu.xlsx", "Şablon"
              )}
            >
              <FileDown size={14} /> Excel Şablonu İndir
            </button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}>
              <Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excel'den İçe Aktar"}
            </button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excel'e Aktar</button>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "span 2" }}>
            <label className="field-label">Cari İsmi</label>
            <input className="input" list="cari-list" placeholder="Firma / tedarikçi adı" value={form.cari} onChange={set("cari")} />
            <datalist id="cari-list">{cariler.map((c) => <option key={c} value={c} />)}</datalist>
          </div>
          <div><label className="field-label">Proje Kodu</label><input className="input" placeholder="Örn: 2026-092" value={form.projeKodu} onChange={set("projeKodu")} /></div>
          <div><label className="field-label">Proje Adı</label><input className="input" placeholder="Örn: ENDERUS" value={form.projeAdi} onChange={set("projeAdi")} /></div>
          <div><label className="field-label">Kalite</label><input className="input" placeholder="Örn: 4140 KALİTE" value={form.kalite} onChange={set("kalite")} /></div>
          <div><label className="field-label">Miktar (Kg)</label><input className="input" type="number" min="0" step="0.01" placeholder="0" value={form.miktar} onChange={set("miktar")} /></div>
          <div>
            <label className="field-label">Durumu</label>
            <input className="input" list="durum-list" placeholder="Seç veya yaz" value={form.durumu} onChange={set("durumu")} />
            <datalist id="durum-list">{DURUM_SECENEKLERI.map((d) => <option key={d} value={d} />)}</datalist>
          </div>
          <div style={{ gridColumn: "span 2" }}><label className="field-label">Açıklama 1 (Ölçü / Adet)</label><input className="input" placeholder="Örn: Ø30X375 1 ADET 10 KALAY" value={form.aciklama1} onChange={set("aciklama1")} /></div>
          <div style={{ gridColumn: "span 2" }}><label className="field-label">Açıklama 2 (Parça)</label><input className="input" placeholder="Parça kodu / adı" value={form.aciklama2} onChange={set("aciklama2")} /></div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button onClick={submit} style={{ display: "flex", alignItems: "center", gap: 7, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
            <Plus size={16} /> {gorunum === "tamamlanan" ? "Tamamlanan Olarak Ekle" : "Açık Sipariş Olarak Ekle"}
          </button>
          {msg && <span style={{ fontSize: 12.5, color: "#8b929a" }}>{msg}</span>}
        </div>
        {iceMsg && (
          <div style={{ marginTop: 14, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{iceMsg}</div>
        )}
        <div style={{ fontSize: 12, color: "#6b7178", marginTop: 10 }}>
          Excel'den içe aktarırken sütun sırası: Cari İsmi, Proje Kodu, Proje Adı, Kalite, Açıklama 1, Açıklama 2, Miktar (Kg), Durumu. Başlık satırı olabilir. İçe aktarılanlar şu an açık olan <b>{gorunum === "tamamlanan" ? "Tamamlanan" : "Açık Siparişler"}</b> sekmesine eklenir.
        </div>
      </div>

      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1", position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Cari, proje, kalite, parça ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Cari</label>
            <select className="input" value={f.cari} onChange={setF2("cari")}>
              <option value="">Tümü</option>
              {cariler.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Proje Kodu</label>
            <select className="input" value={f.projeKodu} onChange={setF2("projeKodu")}>
              <option value="">Tümü</option>
              {projeler.map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Durumu</label>
            <select className="input" value={f.durumu} onChange={setF2("durumu")}>
              <option value="">Tümü</option>
              {DURUM_SECENEKLERI.map((d) => <option key={d} value={d}>{d}</option>)}
            </select>
          </div>
        </div>
      </div>

      {secililer.size > 0 && (
        <div className="card" style={{ padding: "12px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", borderColor: "#2dd4bf" }}>
          <span style={{ fontSize: 13, fontWeight: 700 }}>{secililer.size} kayıt seçili</span>
          <div style={{ display: "flex", gap: 8 }}>
            {gorunum === "acik" ? (
              <button onClick={() => topluIsaretle(true)} style={{ background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "8px 14px", fontWeight: 700, fontSize: 12.5, cursor: "pointer" }}>
                Tamamlandı Olarak İşaretle
              </button>
            ) : (
              <button onClick={() => topluIsaretle(false)} style={{ background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "8px 14px", fontWeight: 700, fontSize: 12.5, cursor: "pointer" }}>
                Açık Siparişe Geri Al
              </button>
            )}
            <button onClick={() => setSecililer(new Set())} className="btn-ghost">Seçimi Temizle</button>
          </div>
        </div>
      )}

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>
          {gorunum === "acik" ? "Açık Siparişler" : "Tamamlanan"} ({filtrelenmis.length})
        </div>
        <div style={{ overflowX: "auto", maxHeight: 560, overflowY: "auto" }}>
          <table>
            <thead>
              <tr>
                <th style={{ width: 36 }}><input type="checkbox" checked={hepsiSecili} onChange={tumunuSecToggle} /></th>
                <th>Cari İsmi</th><th>Proje Kodu</th><th>Proje Adı</th><th>Kalite</th><th>Açıklama 1</th><th>Açıklama 2</th><th>Miktar (Kg)</th><th>Durumu</th><th></th>
              </tr>
            </thead>
            <tbody>
              {filtrelenmis.length === 0 && <tr><td colSpan={10} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>{gorunum === "acik" ? "Açık sipariş yok." : "Henüz tamamlanan kayıt yok."}</td></tr>}
              {filtrelenmis.map((h) => (
                <tr key={h.id}>
                  <td><input type="checkbox" checked={secililer.has(h.id)} onChange={() => birSecToggle(h.id)} /></td>
                  <td>{h.cari}</td>
                  <td style={{ fontFamily: "monospace" }}>{h.projeKodu || "—"}</td>
                  <td>{h.projeAdi || "—"}</td>
                  <td>{h.kalite || "—"}</td>
                  <td>{h.aciklama1 || "—"}</td>
                  <td>{h.aciklama2 || "—"}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#2dd4bf" }}>{(h.miktar || 0).toLocaleString("tr-TR")} kg</td>
                  <td>{h.durumu ? <span className="pill">{h.durumu}</span> : "—"}</td>
                  <td style={{ display: "flex", gap: 6 }}>
                    {gorunum === "acik" ? (
                      <button onClick={() => durumDegistir(h.id, true)} title="Tamamlandı olarak işaretle" style={{ background: "none", border: "1px solid #3d6169", color: "#8b929a", borderRadius: 6, padding: "4px 8px", cursor: "pointer", fontSize: 11 }}>✓ Tamam</button>
                    ) : (
                      <button onClick={() => durumDegistir(h.id, false)} title="Açık siparişe geri al" style={{ background: "none", border: "1px solid #3d6169", color: "#8b929a", borderRadius: 6, padding: "4px 8px", cursor: "pointer", fontSize: 11 }}>↺ Geri Al</button>
                    )}
                    <button onClick={() => sil(h.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Malzeme Tanımları ----------
function MetalMalzemeYonetimi({ metalMalzemeler }) {
  const [ad, setAd] = useState("");
  const [yogunluk, setYogunluk] = useState("");
  const [msg, setMsg] = useState("");

  const ekle = async () => {
    if (!ad.trim() || !yogunluk) {
      setMsg("Malzeme adı ve yoğunluk zorunlu.");
      setTimeout(() => setMsg(""), 2500);
      return;
    }
    if (metalMalzemeler.some((m) => m.ad.toLowerCase() === ad.trim().toLowerCase())) {
      setMsg("Bu malzeme zaten var.");
      setTimeout(() => setMsg(""), 2500);
      return;
    }
    await addDoc(collection(db, "metal_malzemeler"), { ad: ad.trim(), yogunluk: Number(yogunluk) });
    setAd(""); setYogunluk("");
  };

  const sil = async (id) => { await deleteDoc(doc(db, "metal_malzemeler", id)); };

  const varsayilanEkle = async () => {
    const mevcut = new Set(metalMalzemeler.map((m) => m.ad.toLowerCase()));
    const eklenecek = VARSAYILAN_MALZEMELER.filter((m) => !mevcut.has(m.ad.toLowerCase()));
    for (const m of eklenecek) {
      await addDoc(collection(db, "metal_malzemeler"), m);
    }
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}>Yeni Malzeme Ekle</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div><label className="field-label">Malzeme Adı</label><input className="input" placeholder="Örn: Pirinç" value={ad} onChange={(e) => setAd(e.target.value)} /></div>
          <div><label className="field-label">Yoğunluk (g/cm³)</label><input className="input" type="number" step="0.001" placeholder="Örn: 8.40" value={yogunluk} onChange={(e) => setYogunluk(e.target.value)} /></div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button onClick={ekle} style={{ display: "flex", alignItems: "center", gap: 7, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
            <Plus size={16} /> Malzeme Ekle
          </button>
          {msg && <span style={{ fontSize: 12.5, color: "#e07a6b" }}>{msg}</span>}
        </div>
        <div style={{ marginTop: 14 }}>
          <button className="btn-ghost" onClick={varsayilanEkle}>Varsayılan Malzemeleri Ekle (Çelik, Paslanmaz, Bronz, Kestamid, Alüminyum)</button>
        </div>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Malzeme Listesi ({metalMalzemeler.length})</div>
        {metalMalzemeler.length === 0 ? (
          <div style={{ color: "#6b7178", textAlign: "center", padding: 32, fontSize: 13.5 }}>Henüz malzeme tanımlanmadı.</div>
        ) : (
          <table>
            <thead><tr><th>Malzeme Adı</th><th>Yoğunluk (g/cm³)</th><th></th></tr></thead>
            <tbody>
              {metalMalzemeler.map((m) => (
                <tr key={m.id}>
                  <td>{m.ad}</td>
                  <td style={{ fontFamily: "monospace" }}>{m.yogunluk}</td>
                  <td><button onClick={() => sil(m.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ---------- Metal Talep - ölçü/hesaplama ortak formu ----------
function MetalOlcuFormu({ tur, setTur, dims, setDims, yogunluk, setYogunluk, malzemeler, malzemeAdi, setMalzemeAdi, boy, setBoy, boyBirim, setBoyBirim, adet, setAdet }) {
  return (
    <>
      <div style={{ marginBottom: 6, fontSize: 12.5, fontWeight: 600, color: "#8b929a" }}>1. Malzeme</div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
        {malzemeler.map((m) => (
          <button
            key={m.id}
            onClick={() => { setMalzemeAdi(m.ad); setYogunluk(String(m.yogunluk)); }}
            style={{
              padding: "9px 14px", borderRadius: 8, cursor: "pointer", fontWeight: 600, fontSize: 12.5,
              background: malzemeAdi === m.ad ? "#2dd4bf" : "#1b333c",
              color: malzemeAdi === m.ad ? "#142a30" : "#c7cbd1",
              border: `1px solid ${malzemeAdi === m.ad ? "#2dd4bf" : "#2a4b52"}`,
            }}
          >
            {m.ad}
          </button>
        ))}
        {malzemeler.length === 0 && <span style={{ fontSize: 12.5, color: "#6b7178" }}>Önce "Malzeme Tanımları" ekranından malzeme ekle.</span>}
      </div>

      <div style={{ marginBottom: 6, fontSize: 12.5, fontWeight: 600, color: "#8b929a" }}>2. Kesit Türü</div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
        {KESIT_TIPLERI.map((k) => (
          <button
            key={k.id}
            onClick={() => setTur(k.id)}
            style={{
              padding: "9px 14px", borderRadius: 8, cursor: "pointer", fontWeight: 600, fontSize: 12.5,
              background: tur === k.id ? "#142a30" : "#1b333c",
              color: tur === k.id ? "#fff" : "#c7cbd1",
              border: `1px solid ${tur === k.id ? "#142a30" : "#2a4b52"}`,
            }}
          >
            {k.label}
          </button>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 14, marginBottom: 14 }}>
        {OLCU_ALANLARI[tur].map((f) => (
          <div key={f.id}>
            <label className="field-label">{f.label}</label>
            <input
              className="input" type="number" step="0.1" min="0"
              value={dims[f.id] ?? f.def}
              onChange={(e) => setDims((d) => ({ ...d, [f.id]: Number(e.target.value) || 0 }))}
            />
          </div>
        ))}
        <div>
          <label className="field-label">Boy</label>
          <div style={{ display: "flex", gap: 6 }}>
            <input className="input" type="number" step="1" min="0" value={boy} onChange={(e) => setBoy(e.target.value)} />
            <select className="input" style={{ width: 66, flexShrink: 0 }} value={boyBirim} onChange={(e) => setBoyBirim(e.target.value)}>
              <option value="mm">mm</option><option value="cm">cm</option><option value="m">m</option>
            </select>
          </div>
        </div>
        <div><label className="field-label">Adet</label><input className="input" type="number" step="1" min="1" value={adet} onChange={(e) => setAdet(e.target.value)} /></div>
        <div><label className="field-label">Yoğunluk (g/cm³)</label><input className="input" type="number" step="0.001" value={yogunluk} onChange={(e) => setYogunluk(e.target.value)} /></div>
      </div>
    </>
  );
}

// ---------- Metal Talep - Yeni Talep Ekle ----------
function MetalTalepEkle({ metalMalzemeler }) {
  const malzemeListesi = useMemo(() => birlesikMalzemeler(metalMalzemeler), [metalMalzemeler]);
  const [talepNo, setTalepNo] = useState("");
  const [tur, setTur] = useState("mil");
  const [dims, setDims] = useState({});
  const [malzemeAdi, setMalzemeAdi] = useState(malzemeListesi[0]?.ad || "");
  const [yogunluk, setYogunluk] = useState(malzemeListesi[0] ? String(malzemeListesi[0].yogunluk) : "7.85");
  const [boy, setBoy] = useState("500");
  const [boyBirim, setBoyBirim] = useState("mm");
  const [adet, setAdet] = useState("1");
  const [fiyat, setFiyat] = useState("");
  const [msg, setMsg] = useState("");
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  const [iceMsg, setIceMsg] = useState("");
  const dosyaRef = useRef(null);

  useEffect(() => {
    setDims({});
  }, [tur]);

  useEffect(() => {
    if (!malzemeAdi && malzemeListesi[0]) {
      setMalzemeAdi(malzemeListesi[0].ad);
      setYogunluk(String(malzemeListesi[0].yogunluk));
    }
  }, [malzemeListesi]);

  const efektifDims = useMemo(() => {
    const sonuc = {};
    OLCU_ALANLARI[tur].forEach((f) => { sonuc[f.id] = dims[f.id] ?? f.def; });
    return sonuc;
  }, [tur, dims]);

  const boyM = boyMetreCevir(boy, boyBirim);
  const birimKg = kgMetre(tur, efektifDims, Number(yogunluk)) * boyM;
  const toplamKg = birimKg * (Number(adet) || 0);

  const ekle = async () => {
    if (boyM <= 0 || Number(adet) <= 0 || Number(yogunluk) <= 0) {
      setMsg("Boy, adet ve yoğunluk değerlerini kontrol et.");
      setTimeout(() => setMsg(""), 3000);
      return;
    }
    const tutar = fiyat ? toplamKg * Number(fiyat) : 0;
    await addDoc(collection(db, "metal_talepler"), {
      talepNo: talepNo.trim() || "—", tur, malzemeAdi, yogunluk: Number(yogunluk),
      dims: efektifDims, dimLabel: olcuEtiketi(tur, efektifDims),
      boy: boyM, adet: Number(adet), birimKg, toplamKg,
      fiyat: fiyat ? Number(fiyat) : null, tutar,
      tarih: Date.now(),
    });
    setMsg("Talebe eklendi.");
    setTimeout(() => setMsg(""), 2000);
  };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true);
    setIceMsg("");
    try {
      const kayitlar = await excelDenMetalOku(dosya, malzemeListesi);
      if (kayitlar.length === 0) {
        setIceMsg("Dosyada geçerli satır bulunamadı.");
      } else {
        const veriler = kayitlar.map((k) => ({ ...k, tarih: Date.now() }));
        const { basarili, basarisiz } = await guvenliTopluYaz("metal_talepler", veriler, (yapilan, toplam, hatali) => {
          setIceMsg(`${yapilan} / ${toplam} kayıt işlendi${hatali > 0 ? ` (${hatali} tekrar deneniyor)` : ""}…`);
        });
        setIceMsg(`${basarili} kayıt içe aktarıldı${basarisiz > 0 ? `, ${basarisiz} tanesi eklenemedi` : ""}.`);
      }
    } catch (err) {
      console.error(err);
      setIceMsg("Hata: " + (err?.message || "bilinmeyen hata"));
    }
    setIceAktariliyor(false);
    e.target.value = "";
    setTimeout(() => setIceMsg(""), 8000);
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Excel İşlemleri</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button
              className="btn-ghost"
              onClick={() => sablonIndir(
                ["Talep No", "Tür", "Malzeme", "Ölçü1 (mm)", "Ölçü2 (mm)", "Boy (m)", "Adet", "Fiyat (TL/kg)"],
                [["ARC-2026-014", "Mil", "Çelik", 20, "", 6, 10, ""], ["ARC-2026-014", "Lama", "Bronz", 10, 50, 6, 2, ""]],
                "metal-talep-sablonu.xlsx", "Şablon"
              )}
            >
              <FileDown size={14} /> Şablon İndir
            </button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}>
              <Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excelden İçeri Al"}
            </button>
          </div>
        </div>
        {iceMsg && <div style={{ fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{iceMsg}</div>}
        <div style={{ fontSize: 12, color: "#6b7178", marginTop: 10 }}>Sütun sırası: Talep No, Tür (Mil/Kare/Lama/Boru), Malzeme, Ölçü1, Ölçü2 (Lama/Boru için), Boy (m), Adet, Fiyat.</div>
      </div>

      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}>Talep Bilgisi</div>
        <div style={{ marginBottom: 16 }}>
          <label className="field-label">Talep / Proje No</label>
          <input className="input" placeholder="Örn: ARC-2026-014" value={talepNo} onChange={(e) => setTalepNo(e.target.value)} />
        </div>

        <MetalOlcuFormu
          tur={tur} setTur={setTur} dims={dims} setDims={setDims}
          yogunluk={yogunluk} setYogunluk={setYogunluk}
          malzemeler={malzemeListesi} malzemeAdi={malzemeAdi} setMalzemeAdi={setMalzemeAdi}
          boy={boy} setBoy={setBoy} boyBirim={boyBirim} setBoyBirim={setBoyBirim}
          adet={adet} setAdet={setAdet}
        />

        <div><label className="field-label">Birim Fiyat (TL/kg, opsiyonel)</label><input className="input" type="number" step="0.01" placeholder="opsiyonel" value={fiyat} onChange={(e) => setFiyat(e.target.value)} style={{ maxWidth: 220 }} /></div>

        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "#16232a", border: "1px dashed #2a4b52", borderRadius: 8, padding: "12px 16px", margin: "16px 0", flexWrap: "wrap", gap: 8 }}>
          <span style={{ fontFamily: "monospace", fontSize: 13, color: "#8b929a" }}>
            Parça ağırlığı: <b style={{ color: "#2dd4bf" }}>{birimKg > 0 ? birimKg.toFixed(3) : "–"}</b> kg · Toplam: <b style={{ color: "#2dd4bf" }}>{toplamKg > 0 ? toplamKg.toFixed(3) : "–"}</b> kg
          </span>
        </div>

        <button onClick={ekle} style={{ width: "100%", background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "12px 0", fontWeight: 700, fontSize: 14, cursor: "pointer" }}>
          Listeye Ekle
        </button>
        {msg && <div style={{ marginTop: 10, fontSize: 12.5, color: "#2dd4bf" }}>{msg}</div>}
      </div>
    </div>
  );
}

// ---------- Metal Talep - Hızlı KG Hesabı ----------
function MetalHizliHesap({ metalMalzemeler, kullanici }) {
  const malzemeListesi = useMemo(() => birlesikMalzemeler(metalMalzemeler), [metalMalzemeler]);
  const [tur, setTur] = useState("mil");
  const [dims, setDims] = useState({});
  const [malzemeAdi, setMalzemeAdi] = useState(malzemeListesi[0]?.ad || "");
  const [yogunluk, setYogunluk] = useState(malzemeListesi[0] ? String(malzemeListesi[0].yogunluk) : "7.85");
  const [boy, setBoy] = useState("500");
  const [boyBirim, setBoyBirim] = useState("mm");
  const [adet, setAdet] = useState("1");
  const [not, setNot] = useState("");
  const [msg, setMsg] = useState("");
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  const [iceMsg, setIceMsg] = useState("");
  const dosyaRef = useRef(null);

  useEffect(() => { setDims({}); }, [tur]);
  useEffect(() => {
    if (!malzemeAdi && malzemeListesi[0]) {
      setMalzemeAdi(malzemeListesi[0].ad);
      setYogunluk(String(malzemeListesi[0].yogunluk));
    }
  }, [malzemeListesi]);

  const efektifDims = useMemo(() => {
    const sonuc = {};
    OLCU_ALANLARI[tur].forEach((f) => { sonuc[f.id] = dims[f.id] ?? f.def; });
    return sonuc;
  }, [tur, dims]);

  const boyM = boyMetreCevir(boy, boyBirim);
  const birimKg = kgMetre(tur, efektifDims, Number(yogunluk)) * boyM;
  const toplamKg = birimKg * (Number(adet) || 0);

  const kaydet = async () => {
    if (boyM <= 0 || Number(adet) <= 0 || Number(yogunluk) <= 0) {
      setMsg("Boy, adet ve yoğunluk değerlerini kontrol et.");
      setTimeout(() => setMsg(""), 3000);
      return;
    }
    await addDoc(collection(db, "metal_talepler"), {
      talepNo: "—", tur, malzemeAdi, yogunluk: Number(yogunluk),
      dims: efektifDims, dimLabel: olcuEtiketi(tur, efektifDims),
      boy: boyM, adet: Number(adet), birimKg, toplamKg,
      fiyat: null, tutar: 0, not: not.trim(),
      kullanici: kullanici?.email || "—", tarih: Date.now(),
    });
    setNot("");
    setMsg("Ölçüm kaydedildi ✓ — \"Geçmiş Ölçümler\" sekmesinden görebilirsin.");
    setTimeout(() => setMsg(""), 3500);
  };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true); setIceMsg("");
    try {
      const kayitlar = await excelDenMetalOku(dosya, malzemeListesi);
      if (kayitlar.length === 0) {
        setIceMsg("Dosyada geçerli satır bulunamadı.");
      } else {
        const veriler = kayitlar.map((k) => ({ ...k, kullanici: kullanici?.email || "—", tarih: Date.now() }));
        const { basarili, basarisiz } = await guvenliTopluYaz("metal_talepler", veriler);
        setIceMsg(`${basarili} ölçüm eklendi${basarisiz > 0 ? `, ${basarisiz} başarısız` : ""}.`);
      }
    } catch (err) {
      console.error(err);
      setIceMsg("Hata: " + (err?.message || "bilinmeyen hata"));
    }
    setIceAktariliyor(false); e.target.value = ""; setTimeout(() => setIceMsg(""), 6000);
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontWeight: 700, fontSize: 15 }}>Hızlı KG Hesaplama</div>
            <div style={{ fontSize: 12, color: "#6b7178", marginTop: 2 }}>Ölçüleri gir, otomatik hesaplasın. "Kaydet" dediğinde "Geçmiş Ölçümler" sayfasına otomatik eklenir.</div>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn-ghost" onClick={() => sablonIndir(["Talep No", "Tür", "Malzeme", "Ölçü1 (mm)", "Ölçü2 (mm)", "Boy (m)", "Adet", "Fiyat (TL/kg)"], [["—", "Mil", "Çelik", 20, "", 6, 10, ""]], "metal-olcum-sablonu.xlsx", "Şablon")}><FileDown size={14} /> Şablon İndir</button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}><Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excel'den İçe Aktar"}</button>
          </div>
        </div>
        {iceMsg && <div style={{ marginBottom: 14, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{iceMsg}</div>}

        <MetalOlcuFormu
          tur={tur} setTur={setTur} dims={dims} setDims={setDims}
          yogunluk={yogunluk} setYogunluk={setYogunluk}
          malzemeler={malzemeListesi} malzemeAdi={malzemeAdi} setMalzemeAdi={setMalzemeAdi}
          boy={boy} setBoy={setBoy} boyBirim={boyBirim} setBoyBirim={setBoyBirim}
          adet={adet} setAdet={setAdet}
        />

        <div style={{ marginBottom: 14 }}>
          <label className="field-label">Not (opsiyonel)</label>
          <input className="input" placeholder="Örn: hangi parça/proje için" value={not} onChange={(e) => setNot(e.target.value)} />
        </div>

        <div style={{ display: "flex", gap: 24, flexWrap: "wrap", background: "#16232a", border: "1px dashed #2a4b52", borderRadius: 8, padding: 20, marginBottom: 16 }}>
          <div>
            <div style={{ fontSize: 11, color: "#6b7178", textTransform: "uppercase", letterSpacing: "0.06em" }}>Parça Ağırlığı</div>
            <div><b style={{ fontSize: 26, color: "#2dd4bf" }}>{birimKg > 0 ? birimKg.toFixed(3) : "–"}</b> <span style={{ fontSize: 14, color: "#8b929a" }}>kg</span></div>
          </div>
          <div>
            <div style={{ fontSize: 11, color: "#6b7178", textTransform: "uppercase", letterSpacing: "0.06em" }}>Toplam ({adet || 0} adet)</div>
            <div><b style={{ fontSize: 26, color: "#2dd4bf" }}>{toplamKg > 0 ? toplamKg.toFixed(3) : "–"}</b> <span style={{ fontSize: 14, color: "#8b929a" }}>kg</span></div>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <button onClick={kaydet} style={{ display: "flex", alignItems: "center", gap: 7, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
            <Plus size={16} /> Kaydet
          </button>
          {msg && <span style={{ fontSize: 12.5, color: "#8b929a" }}>{msg}</span>}
        </div>
      </div>
    </div>
  );
}

// ---------- Metal Ölçü - Geçmiş Ölçümler ----------
function MetalGecmisOlcumler({ metalTalepler, metalMalzemeler }) {
  const [f, setF] = useState({ arama: "", malzeme: "", tur: "", baslangic: "", bitis: "" });
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const filtrelenmis = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return [...metalTalepler].filter((t) => {
      if (f.malzeme && t.malzemeAdi !== f.malzeme) return false;
      if (f.tur && t.tur !== f.tur) return false;
      if (f.baslangic && t.tarih && new Date(t.tarih).toISOString().slice(0, 10) < f.baslangic) return false;
      if (f.bitis && t.tarih && new Date(t.tarih).toISOString().slice(0, 10) > f.bitis) return false;
      if (q && !((t.malzemeAdi || "").toLowerCase().includes(q) || (t.dimLabel || "").toLowerCase().includes(q) || (t.not || "").toLowerCase().includes(q))) return false;
      return true;
    }).sort((a, b) => (b.tarih || 0) - (a.tarih || 0));
  }, [metalTalepler, f]);

  const sil = async (id) => { await deleteDoc(doc(db, "metal_talepler", id)); };
  const toplamKg = filtrelenmis.reduce((s, t) => s + (Number(t.toplamKg) || 0), 0);
  const disaAktar = () => excelIndir(
    filtrelenmis.map((t) => ({
      Tarih: t.tarih ? new Date(t.tarih).toLocaleString("tr-TR") : "", Malzeme: t.malzemeAdi, Tür: KESIT_ETIKET[t.tur],
      Ölçü: t.dimLabel, "Boy (m)": t.boy?.toFixed(2), Adet: t.adet, "Toplam Kg": t.toplamKg?.toFixed(3), Not: t.not || "", Kullanıcı: t.kullanici || "",
    })), "gecmis-metal-olcumleri.xlsx", "Ölçümler"
  );

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}>Filtrele</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1", position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Malzeme, ölçü, not ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Malzeme</label>
            <select className="input" value={f.malzeme} onChange={setF2("malzeme")}>
              <option value="">Tümü</option>
              {[...new Set(metalTalepler.map((t) => t.malzemeAdi).filter(Boolean))].map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Kesit Türü</label>
            <select className="input" value={f.tur} onChange={setF2("tur")}>
              <option value="">Tümü</option>
              {KESIT_TIPLERI.map((k) => <option key={k.id} value={k.id}>{k.label}</option>)}
            </select>
          </div>
          <div><label className="field-label">Başlangıç</label><input className="input" type="date" value={f.baslangic} onChange={setF2("baslangic")} /></div>
          <div><label className="field-label">Bitiş</label><input className="input" type="date" value={f.bitis} onChange={setF2("bitis")} /></div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Kayıt Sayısı" value={filtrelenmis.length} />
        <Stat label="Toplam Ağırlık" value={`${toplamKg.toLocaleString("tr-TR", { maximumFractionDigits: 2 })} kg`} highlight />
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>Geçmiş Ölçümler ({filtrelenmis.length})</div>
          <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excele Aktar</button>
        </div>
        <div style={{ overflowX: "auto", maxHeight: 640, overflowY: "auto" }}>
          <table>
            <thead><tr><th>Tarih</th><th>Malzeme</th><th>Tür</th><th>Ölçü</th><th>Boy</th><th>Adet</th><th>Toplam Kg</th><th>Not</th><th></th></tr></thead>
            <tbody>
              {filtrelenmis.length === 0 && <tr><td colSpan={9} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Henüz ölçüm kaydedilmedi.</td></tr>}
              {filtrelenmis.map((t) => (
                <tr key={t.id}>
                  <td style={{ fontFamily: "monospace", whiteSpace: "nowrap" }}>{t.tarih ? new Date(t.tarih).toLocaleString("tr-TR") : "—"}</td>
                  <td>{t.malzemeAdi}</td>
                  <td>{KESIT_ETIKET[t.tur]}</td>
                  <td>{t.dimLabel}</td>
                  <td style={{ fontFamily: "monospace" }}>{t.boy?.toFixed(2)} m</td>
                  <td style={{ fontFamily: "monospace" }}>{t.adet}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#2dd4bf" }}>{t.toplamKg?.toFixed(3)}</td>
                  <td style={{ fontSize: 12.5 }}>{t.not || "—"}</td>
                  <td><button onClick={() => sil(t.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}


// ---------- Metal Talep - Talep Listesi (KULLANILMIYOR - eski) ----------
function MetalTalepListesi({ metalTalepler, metalMalzemeler }) {
  const [f, setF] = useState({ arama: "", tur: "", malzeme: "" });
  const [genisletilen, setGenisletilen] = useState(new Set());
  const [tasiniyor, setTasiniyor] = useState(false);
  const [tasimaMsg, setTasimaMsg] = useState("");
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const filtrelenmis = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return metalTalepler.filter((it) => {
      if (f.tur && it.tur !== f.tur) return false;
      if (f.malzeme && it.malzemeAdi !== f.malzeme) return false;
      if (q && !(
        (it.talepNo || "").toLowerCase().includes(q) ||
        (it.malzemeAdi || "").toLowerCase().includes(q) ||
        (it.dimLabel || "").toLowerCase().includes(q)
      )) return false;
      return true;
    }).sort((a, b) => (b.tarih || 0) - (a.tarih || 0));
  }, [metalTalepler, f]);

  const gruplar = useMemo(() => {
    const map = new Map();
    const sira = [];
    filtrelenmis.forEach((it) => {
      const key = it.talepNo || "—";
      if (!map.has(key)) { map.set(key, []); sira.push(key); }
      map.get(key).push(it);
    });
    return sira.map((key) => ({ key, kalemler: map.get(key) }));
  }, [filtrelenmis]);

  const grupToggle = (key) => {
    setGenisletilen((s) => { const y = new Set(s); if (y.has(key)) y.delete(key); else y.add(key); return y; });
  };

  const sil = async (id) => { await deleteDoc(doc(db, "metal_talepler", id)); };

  const tumunuTemizle = async () => {
    if (metalTalepler.length === 0) return;
    if (!window.confirm("Tüm talep listesi silinsin mi? Bu işlem geri alınamaz.")) return;
    const idler = metalTalepler.map((t) => t.id);
    for (let i = 0; i < idler.length; i += 400) {
      const dilim = idler.slice(i, i + 400);
      const batch = writeBatch(db);
      dilim.forEach((id) => batch.delete(doc(db, "metal_talepler", id)));
      await batch.commit();
    }
  };

  const disaAktar = () => {
    excelIndir(
      filtrelenmis.map((it) => ({
        "Talep": it.talepNo, "Malzeme": it.malzemeAdi, "Tür": KESIT_ETIKET[it.tur],
        "Ölçü": it.dimLabel, "Boy (m)": it.boy, "Adet": it.adet,
        "Birim Kg": it.birimKg?.toFixed(3), "Toplam Kg": it.toplamKg?.toFixed(3),
        "TL/Kg": it.fiyat || "", "Tutar": it.tutar || "",
      })),
      "metal-talep-listesi.xlsx", "Talep Listesi"
    );
  };

  const metinKopyala = () => {
    if (filtrelenmis.length === 0) return;
    let text = "Çelik / Metal Talep Listesi\n" + "-".repeat(55) + "\n";
    filtrelenmis.forEach((it, i) => {
      text += `${i + 1}. [${it.talepNo}] ${it.malzemeAdi} ${KESIT_ETIKET[it.tur]} ${it.dimLabel} — Boy: ${it.boy?.toFixed(2)}m x ${it.adet} adet = ${it.toplamKg?.toFixed(3)} kg`;
      if (it.fiyat) text += ` — ${it.fiyat} TL/kg = ${it.tutar?.toFixed(2)} TL`;
      text += "\n";
    });
    const totKg = filtrelenmis.reduce((s, i) => s + (i.toplamKg || 0), 0);
    const totTutar = filtrelenmis.reduce((s, i) => s + (i.tutar || 0), 0);
    text += "-".repeat(55) + `\nTOPLAM AĞIRLIK: ${totKg.toFixed(2)} kg\n`;
    if (totTutar > 0) text += `TOPLAM TUTAR: ${totTutar.toFixed(2)} TL\n`;
    navigator.clipboard.writeText(text).catch(() => {});
  };

  const eskiVeriyiTasi = async () => {
    if (!window.confirm("Eski SAKLAZ-METALERP programındaki tüm talep ve malzeme verileri bu programa kopyalanacak. Devam edilsin mi?")) return;
    setTasiniyor(true);
    setTasimaMsg("Eski veriler okunuyor…");
    try {
      const eskiDb = eskiMetalErpDb();
      const snap = await getDoc(doc(eskiDb, "metalerp", "shared-data"));
      if (!snap.exists()) {
        setTasimaMsg("Eski programda veri bulunamadı.");
        setTasiniyor(false);
        return;
      }
      const veri = snap.data();
      const eskiMalzemeler = veri.materials || {};
      const eskiItems = veri.items || [];

      // Malzemeleri taşı (aynı isimde varsa atla)
      const mevcutAdlar = new Set(metalMalzemeler.map((m) => m.ad.toLowerCase()));
      const yeniMalzemeler = Object.values(eskiMalzemeler)
        .filter((m) => m && m.label && !mevcutAdlar.has(m.label.toLowerCase()))
        .map((m) => ({ ad: m.label, yogunluk: m.density }));
      if (yeniMalzemeler.length) {
        await guvenliTopluYaz("metal_malzemeler", yeniMalzemeler);
      }

      // Talepleri taşı - zaten taşınmış olanları (eskiId ile) tekrar eklemesin
      const mevcutEskiIdler = new Set(metalTalepler.map((t) => t.eskiId).filter(Boolean));
      const yeniTalepler = eskiItems
        .filter((it) => !mevcutEskiIdler.has(it.id))
        .map((it) => {
          const dims = it.type === "mil" ? { cap: it.dims?.cap }
            : it.type === "kare" ? { kenar: it.dims?.kenar }
            : it.type === "lama" ? { kalinlik: it.dims?.kalinlik, genislik: it.dims?.genislik }
            : { disCap: it.dims?.disCap, etKalinligi: it.dims?.etKalinligi };
          return {
            eskiId: it.id, talepNo: it.reqName || "—", tur: it.type, malzemeAdi: it.matLabel,
            yogunluk: it.density, dims, dimLabel: it.dimLabel,
            boy: it.boy, adet: it.adet, birimKg: it.unitKg, toplamKg: it.totalKg,
            fiyat: it.fiyat || null, tutar: it.tutar || 0,
            tarih: it.date ? new Date(it.date).getTime() : Date.now(),
          };
        });

      if (yeniTalepler.length === 0) {
        setTasimaMsg("Taşınacak yeni kayıt yok (hepsi zaten taşınmış).");
      } else {
        setTasimaMsg(`${yeniTalepler.length} kayıt taşınıyor…`);
        const { basarili, basarisiz } = await guvenliTopluYaz("metal_talepler", yeniTalepler, (yapilan, toplam) => {
          setTasimaMsg(`${yapilan} / ${toplam} kayıt taşınıyor…`);
        });
        setTasimaMsg(`${basarili} kayıt taşındı${basarisiz > 0 ? `, ${basarisiz} tanesi başarısız` : ""}.`);
      }
    } catch (err) {
      console.error(err);
      setTasimaMsg("Taşıma sırasında hata oluştu: " + (err?.message || "bilinmeyen hata") + ". Eski projeye erişim engellenmiş olabilir.");
    }
    setTasiniyor(false);
    setTimeout(() => setTasimaMsg(""), 10000);
  };

  const toplamKg = filtrelenmis.reduce((s, i) => s + (i.toplamKg || 0), 0);
  const toplamTutar = filtrelenmis.reduce((s, i) => s + (i.tutar || 0), 0);
  const malzemeAdlari = [...new Set(metalTalepler.map((t) => t.malzemeAdi).filter(Boolean))];

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Filtrele</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn-ghost" onClick={eskiVeriyiTasi} disabled={tasiniyor}>
              <RefreshCw size={14} /> {tasiniyor ? "Taşınıyor…" : "Eski Metal-Erp Verilerini İçe Aktar"}
            </button>
            <button className="btn-ghost" onClick={metinKopyala}><Copy size={14} /> Metin Olarak Kopyala</button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excele Aktar</button>
          </div>
        </div>
        {tasimaMsg && <div style={{ marginBottom: 12, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{tasimaMsg}</div>}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1", position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Talep no, malzeme veya ölçü ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Tür</label>
            <select className="input" value={f.tur} onChange={setF2("tur")}>
              <option value="">Tüm Türler</option>
              {KESIT_TIPLERI.map((k) => <option key={k.id} value={k.id}>{k.label}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Malzeme</label>
            <select className="input" value={f.malzeme} onChange={setF2("malzeme")}>
              <option value="">Tüm Malzemeler</option>
              {malzemeAdlari.map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Görüntülenen Kayıt" value={filtrelenmis.length} />
        <Stat label="Toplam Ağırlık" value={`${toplamKg.toLocaleString("tr-TR", { maximumFractionDigits: 2 })} kg`} highlight />
        <Stat label="Toplam Tutar" value={`${toplamTutar.toLocaleString("tr-TR", { maximumFractionDigits: 2 })} TL`} />
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>Talep Listesi ({gruplar.length} talep)</div>
          <button onClick={tumunuTemizle} className="btn-ghost" style={{ color: "#e07a6b", borderColor: "#5a2a2a" }}>Tüm Listeyi Temizle</button>
        </div>
        <div style={{ overflowX: "auto", maxHeight: 640, overflowY: "auto" }}>
          <table>
            <thead><tr><th></th><th>Talep</th><th>Malzeme</th><th>Tür</th><th>Ölçü</th><th>Boy</th><th>Adet</th><th>Birim Kg</th><th>Toplam Kg</th><th>TL/Kg</th><th>Tutar</th><th></th></tr></thead>
            <tbody>
              {gruplar.length === 0 && <tr><td colSpan={12} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Kayıt bulunamadı.</td></tr>}
              {gruplar.map(({ key, kalemler }) => {
                const acik = genisletilen.has(key);
                const kg = kalemler.reduce((s, i) => s + (i.toplamKg || 0), 0);
                const tutar = kalemler.reduce((s, i) => s + (i.tutar || 0), 0);
                return (
                  <React.Fragment key={key}>
                    <tr onClick={() => grupToggle(key)} style={{ cursor: "pointer", background: "#16232a" }}>
                      <td>{acik ? <ChevronDown size={14} /> : <ChevronRight size={14} />}</td>
                      <td colSpan={5}><span className="pill">{key}</span> <span style={{ color: "#6b7178", fontSize: 12 }}>{kalemler.length} kalem</span></td>
                      <td colSpan={3} style={{ fontFamily: "monospace" }}>{kg.toFixed(2)} kg</td>
                      <td colSpan={3} style={{ fontFamily: "monospace", color: "#2dd4bf", fontWeight: 700 }}>{tutar ? tutar.toFixed(2) + " TL" : "—"}</td>
                    </tr>
                    {acik && kalemler.map((it) => (
                      <tr key={it.id}>
                        <td></td>
                        <td></td>
                        <td>{it.malzemeAdi}</td>
                        <td>{KESIT_ETIKET[it.tur]}</td>
                        <td>{it.dimLabel}</td>
                        <td style={{ fontFamily: "monospace" }}>{it.boy?.toFixed(2)} m</td>
                        <td style={{ fontFamily: "monospace" }}>{it.adet}</td>
                        <td style={{ fontFamily: "monospace" }}>{it.birimKg?.toFixed(3)}</td>
                        <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#2dd4bf" }}>{it.toplamKg?.toFixed(3)}</td>
                        <td style={{ fontFamily: "monospace" }}>{it.fiyat || "—"}</td>
                        <td style={{ fontFamily: "monospace" }}>{it.fiyat ? it.tutar?.toFixed(2) + " TL" : "—"}</td>
                        <td><button onClick={() => sil(it.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                      </tr>
                    ))}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Stok Kayıtları Sil ----------
function StokSilme({ records }) {
  const [f, setF] = useState({ arama: "", takim: "", makine: "" });
  const [secililer, setSecililer] = useState(new Set());
  const [durum, setDurum] = useState("");
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const takimlar = [...new Set(records.map((r) => r.takim).filter(Boolean))];
  const makineler = [...new Set(records.map((r) => r.makine).filter(Boolean))];

  const filtrelenmis = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return records.filter((r) => {
      if (f.takim && r.takim !== f.takim) return false;
      if (f.makine && r.makine !== f.makine) return false;
      if (q && !(
        (r.takim || "").toLowerCase().includes(q) ||
        (r.magaza || "").toLowerCase().includes(q) ||
        (r.makine || "").toLowerCase().includes(q) ||
        (r.urun || "").toLowerCase().includes(q) ||
        (r.tarih || "").includes(q)
      )) return false;
      return true;
    }).sort((a, b) => (b.olusturma || 0) - (a.olusturma || 0));
  }, [records, f]);

  const hepsiSecili = filtrelenmis.length > 0 && filtrelenmis.every((r) => secililer.has(r.id));
  const tumunuSecToggle = () => setSecililer(hepsiSecili ? new Set() : new Set(filtrelenmis.map((r) => r.id)));
  const birSecToggle = (id) => setSecililer((s) => { const y = new Set(s); if (y.has(id)) y.delete(id); else y.add(id); return y; });

  const secilenleriSil = async () => {
    if (secililer.size === 0) return;
    if (!window.confirm(`${secililer.size} stok kaydı kalıcı olarak silinecek. Bu işlem geri alınamaz. Emin misiniz?`)) return;
    setDurum("Siliniyor…");
    const idler = [...secililer];
    for (let i = 0; i < idler.length; i += 400) {
      const dilim = idler.slice(i, i + 400);
      const batch = writeBatch(db);
      dilim.forEach((id) => batch.delete(doc(db, "records", id)));
      await batch.commit();
    }
    setSecililer(new Set());
    setDurum(`${idler.length} kayıt silindi.`);
    setTimeout(() => setDurum(""), 4000);
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}>Stok Kayıtlarını Filtrele</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1", position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Takım, mağaza, makine, tarih ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Takım</label>
            <select className="input" value={f.takim} onChange={setF2("takim")}>
              <option value="">Tümü</option>
              {takimlar.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Makine</label>
            <select className="input" value={f.makine} onChange={setF2("makine")}>
              <option value="">Tümü</option>
              {makineler.map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
          </div>
        </div>
      </div>

      {(secililer.size > 0 || durum) && (
        <div className="card" style={{ padding: "12px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", borderColor: "#c0392b" }}>
          <span style={{ fontSize: 13, fontWeight: 700 }}>{durum || `${secililer.size} kayıt seçili`}</span>
          {secililer.size > 0 && !durum && (
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={secilenleriSil} style={{ background: "#c0392b", color: "#fff", border: "none", borderRadius: 7, padding: "8px 14px", fontWeight: 700, fontSize: 12.5, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
                <Trash2 size={14} /> Seçilenleri Sil
              </button>
              <button onClick={() => setSecililer(new Set())} className="btn-ghost">Seçimi Temizle</button>
            </div>
          )}
        </div>
      )}

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Stok Kayıtları ({filtrelenmis.length})</div>
        <div style={{ overflowX: "auto", maxHeight: 560, overflowY: "auto" }}>
          <table>
            <thead>
              <tr>
                <th style={{ width: 36 }}><input type="checkbox" checked={hepsiSecili} onChange={tumunuSecToggle} /></th>
                <th>Tarih</th><th>Takım</th><th>Mağaza</th><th>Makine</th><th>Ürün</th><th>Adet</th>
              </tr>
            </thead>
            <tbody>
              {filtrelenmis.length === 0 && <tr><td colSpan={7} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Kayıt bulunamadı.</td></tr>}
              {filtrelenmis.map((r) => (
                <tr key={r.id}>
                  <td><input type="checkbox" checked={secililer.has(r.id)} onChange={() => birSecToggle(r.id)} /></td>
                  <td style={{ fontFamily: "monospace" }}>{r.tarih}</td>
                  <td>{r.takim}</td>
                  <td>{r.magaza || "—"}</td>
                  <td>{r.makine}</td>
                  <td>{r.urun || "—"}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#2dd4bf" }}>{r.adet}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Hammadde Kayıtları Sil ----------
function HammaddeSilme({ hammaddeler }) {
  const [f, setF] = useState({ arama: "", cari: "", durumu: "" });
  const [secililer, setSecililer] = useState(new Set());
  const [durum, setDurum] = useState("");
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const cariler = [...new Set(hammaddeler.map((h) => h.cari).filter(Boolean))];

  const filtrelenmis = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return hammaddeler.filter((h) => {
      if (f.cari && h.cari !== f.cari) return false;
      if (f.durumu === "acik" && h.tamamlandi) return false;
      if (f.durumu === "tamamlanan" && !h.tamamlandi) return false;
      if (q && !(
        (h.cari || "").toLowerCase().includes(q) ||
        (h.projeKodu || "").toLowerCase().includes(q) ||
        (h.kalite || "").toLowerCase().includes(q) ||
        (h.aciklama2 || "").toLowerCase().includes(q)
      )) return false;
      return true;
    }).sort((a, b) => (b.olusturma || 0) - (a.olusturma || 0));
  }, [hammaddeler, f]);

  const hepsiSecili = filtrelenmis.length > 0 && filtrelenmis.every((h) => secililer.has(h.id));
  const tumunuSecToggle = () => setSecililer(hepsiSecili ? new Set() : new Set(filtrelenmis.map((h) => h.id)));
  const birSecToggle = (id) => setSecililer((s) => { const y = new Set(s); if (y.has(id)) y.delete(id); else y.add(id); return y; });

  const secilenleriSil = async () => {
    if (secililer.size === 0) return;
    if (!window.confirm(`${secililer.size} hammadde kaydı kalıcı olarak silinecek. Bu işlem geri alınamaz. Emin misiniz?`)) return;
    setDurum("Siliniyor…");
    const idler = [...secililer];
    for (let i = 0; i < idler.length; i += 400) {
      const dilim = idler.slice(i, i + 400);
      const batch = writeBatch(db);
      dilim.forEach((id) => batch.delete(doc(db, "hammadde", id)));
      await batch.commit();
    }
    setSecililer(new Set());
    setDurum(`${idler.length} kayıt silindi.`);
    setTimeout(() => setDurum(""), 4000);
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}>Hammadde Kayıtlarını Filtrele</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1", position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Cari, proje, kalite, parça ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Cari</label>
            <select className="input" value={f.cari} onChange={setF2("cari")}>
              <option value="">Tümü</option>
              {cariler.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Durum</label>
            <select className="input" value={f.durumu} onChange={setF2("durumu")}>
              <option value="">Tümü</option>
              <option value="acik">Açık Siparişler</option>
              <option value="tamamlanan">Tamamlanan</option>
            </select>
          </div>
        </div>
      </div>

      {(secililer.size > 0 || durum) && (
        <div className="card" style={{ padding: "12px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", borderColor: "#c0392b" }}>
          <span style={{ fontSize: 13, fontWeight: 700 }}>{durum || `${secililer.size} kayıt seçili`}</span>
          {secililer.size > 0 && !durum && (
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={secilenleriSil} style={{ background: "#c0392b", color: "#fff", border: "none", borderRadius: 7, padding: "8px 14px", fontWeight: 700, fontSize: 12.5, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
                <Trash2 size={14} /> Seçilenleri Sil
              </button>
              <button onClick={() => setSecililer(new Set())} className="btn-ghost">Seçimi Temizle</button>
            </div>
          )}
        </div>
      )}

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Hammadde Kayıtları ({filtrelenmis.length})</div>
        <div style={{ overflowX: "auto", maxHeight: 560, overflowY: "auto" }}>
          <table>
            <thead>
              <tr>
                <th style={{ width: 36 }}><input type="checkbox" checked={hepsiSecili} onChange={tumunuSecToggle} /></th>
                <th>Cari İsmi</th><th>Proje Kodu</th><th>Kalite</th><th>Açıklama 2</th><th>Miktar (Kg)</th><th>Durum</th>
              </tr>
            </thead>
            <tbody>
              {filtrelenmis.length === 0 && <tr><td colSpan={7} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Kayıt bulunamadı.</td></tr>}
              {filtrelenmis.map((h) => (
                <tr key={h.id}>
                  <td><input type="checkbox" checked={secililer.has(h.id)} onChange={() => birSecToggle(h.id)} /></td>
                  <td>{h.cari}</td>
                  <td style={{ fontFamily: "monospace" }}>{h.projeKodu || "—"}</td>
                  <td>{h.kalite || "—"}</td>
                  <td>{h.aciklama2 || "—"}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#2dd4bf" }}>{(h.miktar || 0).toLocaleString("tr-TR")} kg</td>
                  <td>{h.tamamlandi ? <span className="pill">Tamamlandı</span> : <span className="pill" style={{ background: "#1f2d3a", color: "#7fb0e0", borderColor: "#2c4a63" }}>Açık</span>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Depo Stok Sil ----------
function DepoSilme({ depoStok }) {
  const [arama, setArama] = useState("");
  const [secililer, setSecililer] = useState(new Set());
  const [durum, setDurum] = useState("");

  const filtrelenmis = useMemo(() => {
    if (!arama.trim()) return depoStok;
    const q = arama.trim().toLowerCase();
    return depoStok.filter((s) => s.stokKodu.toLowerCase().includes(q) || s.stokAdi.toLowerCase().includes(q));
  }, [depoStok, arama]);

  const hepsiSecili = filtrelenmis.length > 0 && filtrelenmis.every((s) => secililer.has(s.id));
  const tumunuSecToggle = () => setSecililer(hepsiSecili ? new Set() : new Set(filtrelenmis.map((s) => s.id)));
  const birSecToggle = (id) => setSecililer((s) => { const y = new Set(s); if (y.has(id)) y.delete(id); else y.add(id); return y; });

  const secilenleriSil = async () => {
    if (secililer.size === 0) return;
    if (!window.confirm(`${secililer.size} stok kartı kalıcı olarak silinecek (hareket geçmişleri silinmez). Emin misiniz?`)) return;
    setDurum("Siliniyor…");
    const idler = [...secililer];
    for (let i = 0; i < idler.length; i += 400) {
      const dilim = idler.slice(i, i + 400);
      const batch = writeBatch(db);
      dilim.forEach((id) => batch.delete(doc(db, "depo_stok", id)));
      await batch.commit();
    }
    setSecililer(new Set());
    setDurum(`${idler.length} kayıt silindi.`);
    setTimeout(() => setDurum(""), 4000);
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}>Depo Stok Kartlarını Filtrele</div>
        <div style={{ position: "relative" }}>
          <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
          <input className="input" style={{ paddingLeft: 30 }} placeholder="Stok kodu veya adı ara…" value={arama} onChange={(e) => setArama(e.target.value)} />
        </div>
      </div>

      {(secililer.size > 0 || durum) && (
        <div className="card" style={{ padding: "12px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", borderColor: "#c0392b" }}>
          <span style={{ fontSize: 13, fontWeight: 700 }}>{durum || `${secililer.size} kayıt seçili`}</span>
          {secililer.size > 0 && !durum && (
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={secilenleriSil} style={{ background: "#c0392b", color: "#fff", border: "none", borderRadius: 7, padding: "8px 14px", fontWeight: 700, fontSize: 12.5, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
                <Trash2 size={14} /> Seçilenleri Sil
              </button>
              <button onClick={() => setSecililer(new Set())} className="btn-ghost">Seçimi Temizle</button>
            </div>
          )}
        </div>
      )}

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Depo Stok Kartları ({filtrelenmis.length})</div>
        <div style={{ overflowX: "auto", maxHeight: 560, overflowY: "auto" }}>
          <table>
            <thead>
              <tr>
                <th style={{ width: 36 }}><input type="checkbox" checked={hepsiSecili} onChange={tumunuSecToggle} /></th>
                <th>Stok Kodu</th><th>Stok Adı</th><th>Miktar</th><th>Birim</th>
              </tr>
            </thead>
            <tbody>
              {filtrelenmis.length === 0 && <tr><td colSpan={5} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Kayıt bulunamadı.</td></tr>}
              {filtrelenmis.map((s) => (
                <tr key={s.id}>
                  <td><input type="checkbox" checked={secililer.has(s.id)} onChange={() => birSecToggle(s.id)} /></td>
                  <td style={{ fontFamily: "monospace" }}>{s.stokKodu}</td>
                  <td>{s.stokAdi}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#2dd4bf" }}>{s.miktar}</td>
                  <td>{s.birim || "Adet"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ================= FASON TAKİP =================

// ---------- Fason Özet ----------
function FasonOzet({ fasonFirmalar, fasonIsler, fasonHareketler, fasonHatirlaticilar }) {
  const aktifIsSayisi = fasonIsler.filter((j) => j.durum !== "tamamlandi").length;
  const buAy = new Date().toISOString().slice(0, 7);
  const buAyHareketler = fasonHareketler.filter((m) => (m.tarih || "").slice(0, 7) === buAy);
  const gidenAy = buAyHareketler.filter((m) => m.tip === "giden").reduce((s, m) => s + (Number(m.miktar) || 0) * (Number(m.birimFiyat) || 0), 0);
  const gelenAy = buAyHareketler.filter((m) => m.tip === "gelen").reduce((s, m) => s + (Number(m.miktar) || 0) * (Number(m.birimFiyat) || 0), 0);

  const yaklasanlar = [...fasonHatirlaticilar].filter((r) => !r.tamamlandi).sort((a, b) => (a.tarih || "").localeCompare(b.tarih || "")).slice(0, 5);

  const firmaBakiye = (firmaId) => {
    const isIdler = new Set(fasonIsler.filter((j) => j.firmaId === firmaId).map((j) => j.id));
    let giden = 0, gelen = 0;
    fasonHareketler.forEach((m) => {
      if (!isIdler.has(m.isId)) return;
      const t = (Number(m.miktar) || 0) * (Number(m.birimFiyat) || 0);
      if (m.tip === "giden") giden += t; else gelen += t;
    });
    return { giden, gelen, bakiye: giden - gelen };
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Toplam Firma" value={fasonFirmalar.length} />
        <Stat label="Aktif İş" value={aktifIsSayisi} highlight />
        <Stat label="Bu Ay Giden (Hammadde)" value={paraTR(gidenAy)} />
        <Stat label="Bu Ay Gelen (Fason)" value={paraTR(gelenAy)} />
      </div>

      {yaklasanlar.length > 0 && (
        <div className="card" style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Yaklaşan Hatırlatıcılar</div>
          {yaklasanlar.map((r) => {
            const bugun = todayISO();
            const gecikti = r.tarih && r.tarih < bugun;
            const bugunMu = r.tarih === bugun;
            return (
              <div key={r.id} style={{ padding: "12px 20px", borderBottom: "1px solid #223b42", display: "flex", alignItems: "center", gap: 10 }}>
                <Bell size={14} color="#8b929a" />
                <div>
                  <div style={{ fontSize: 13.5 }}>{r.baslik}</div>
                  <div style={{ fontSize: 11.5, color: gecikti ? "#e07a6b" : bugunMu ? "#e8a33d" : "#6b7178" }}>
                    {r.tarih}{gecikti ? " · gecikti" : bugunMu ? " · bugün" : ""}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Firma Bazlı Bakiye</div>
        {fasonFirmalar.length === 0 ? (
          <div style={{ color: "#6b7178", textAlign: "center", padding: 32, fontSize: 13.5 }}>Henüz firma eklenmedi.</div>
        ) : (
          <table>
            <thead><tr><th>Firma</th><th>Giden</th><th>Gelen</th><th>Bakiye</th></tr></thead>
            <tbody>
              {fasonFirmalar.map((f) => {
                const b = firmaBakiye(f.id);
                return (
                  <tr key={f.id}>
                    <td>{f.ad}</td>
                    <td style={{ fontFamily: "monospace", color: "#e8a33d" }}>{paraTR(b.giden)}</td>
                    <td style={{ fontFamily: "monospace", color: "#4b8f5e" }}>{paraTR(b.gelen)}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: b.bakiye >= 0 ? "#2dd4bf" : "#e07a6b" }}>{paraTR(b.bakiye)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ---------- Fason Firmalar ----------
function FasonFirmalar({ fasonFirmalar, fasonIsler, fasonHareketler }) {
  const [form, setForm] = useState({ ad: "", yetkili: "", not: "" });
  const [arama, setArama] = useState("");
  const [msg, setMsg] = useState("");
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  const [iceMsg, setIceMsg] = useState("");
  const dosyaRef = useRef(null);

  const ekle = async () => {
    if (!form.ad.trim()) { setMsg("Firma adı zorunlu."); setTimeout(() => setMsg(""), 2500); return; }
    await addDoc(collection(db, "fason_firmalar"), { ...form });
    setForm({ ad: "", yetkili: "", not: "" });
    setMsg("Firma eklendi."); setTimeout(() => setMsg(""), 2000);
  };
  const sil = async (id) => {
    if (!window.confirm("Bu firma silinecek (işleri ve hareketleri silinmez, ama firma bağlantısı kopar). Emin misiniz?")) return;
    await deleteDoc(doc(db, "fason_firmalar", id));
  };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true); setIceMsg("");
    try {
      const kayitlar = await excelDenFasonFirmaOku(dosya);
      const mevcut = new Set(fasonFirmalar.map((f) => f.ad.toLowerCase()));
      const yeniler = kayitlar.filter((k) => !mevcut.has(k.ad.toLowerCase()));
      const { basarili, basarisiz } = await guvenliTopluYaz("fason_firmalar", yeniler);
      setIceMsg(`${basarili} firma eklendi${basarisiz > 0 ? `, ${basarisiz} başarısız` : ""}.`);
    } catch (err) { console.error(err); setIceMsg("Hata: " + (err?.message || "bilinmeyen hata")); }
    setIceAktariliyor(false); e.target.value = ""; setTimeout(() => setIceMsg(""), 6000);
  };
  const disaAktar = () => excelIndir(fasonFirmalar.map((f) => ({ "Firma Adı": f.ad, "Yetkili": f.yetkili, "Not": f.not })), "fason-firmalar.xlsx", "Firmalar");

  const firmaBakiye = (firmaId) => {
    const isIdler = new Set(fasonIsler.filter((j) => j.firmaId === firmaId).map((j) => j.id));
    let bakiye = 0;
    fasonHareketler.forEach((m) => {
      if (!isIdler.has(m.isId)) return;
      const t = (Number(m.miktar) || 0) * (Number(m.birimFiyat) || 0);
      bakiye += m.tip === "giden" ? t : -t;
    });
    return bakiye;
  };

  const filtrelenmis = useMemo(() => {
    if (!arama.trim()) return fasonFirmalar;
    const q = arama.trim().toLowerCase();
    return fasonFirmalar.filter((f) => f.ad.toLowerCase().includes(q) || (f.yetkili || "").toLowerCase().includes(q));
  }, [fasonFirmalar, arama]);

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Yeni Fason Firma Ekle</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn-ghost" onClick={() => sablonIndir(["FİRMA ADI", "YETKİLİ", "NOT"], [["Örnek Fason Ltd.", "Ahmet Yılmaz - 0532 000 00 00", ""]], "fason-firma-sablonu.xlsx", "Şablon")}><FileDown size={14} /> Şablon İndir</button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}><Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excelden İçeri Al"}</button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excele Aktar</button>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div><label className="field-label">Firma Adı</label><input className="input" value={form.ad} onChange={(e) => setForm((s) => ({ ...s, ad: e.target.value }))} /></div>
          <div><label className="field-label">Yetkili / Telefon</label><input className="input" value={form.yetkili} onChange={(e) => setForm((s) => ({ ...s, yetkili: e.target.value }))} /></div>
          <div style={{ gridColumn: "span 2" }}><label className="field-label">Not</label><input className="input" value={form.not} onChange={(e) => setForm((s) => ({ ...s, not: e.target.value }))} /></div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button onClick={ekle} style={{ display: "flex", alignItems: "center", gap: 7, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}><Plus size={16} /> Firma Ekle</button>
          {msg && <span style={{ fontSize: 12.5, color: "#8b929a" }}>{msg}</span>}
        </div>
        {iceMsg && <div style={{ marginTop: 10, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{iceMsg}</div>}
      </div>

      <div className="card" style={{ padding: "12px 16px" }}>
        <div style={{ position: "relative" }}>
          <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
          <input className="input" style={{ paddingLeft: 30 }} placeholder="Firma ara…" value={arama} onChange={(e) => setArama(e.target.value)} />
        </div>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Firmalar ({filtrelenmis.length})</div>
        {filtrelenmis.length === 0 ? (
          <div style={{ color: "#6b7178", textAlign: "center", padding: 32, fontSize: 13.5 }}>Firma bulunamadı.</div>
        ) : (
          <table>
            <thead><tr><th>Firma Adı</th><th>Yetkili</th><th>İş Sayısı</th><th>Bakiye</th><th></th></tr></thead>
            <tbody>
              {filtrelenmis.map((f) => (
                <tr key={f.id}>
                  <td>{f.ad}</td>
                  <td>{f.yetkili || "—"}</td>
                  <td style={{ fontFamily: "monospace" }}>{fasonIsler.filter((j) => j.firmaId === f.id).length}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: firmaBakiye(f.id) >= 0 ? "#2dd4bf" : "#e07a6b" }}>{paraTR(firmaBakiye(f.id))}</td>
                  <td><button onClick={() => sil(f.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ---------- Fason İşler ----------
function FasonIsler({ fasonFirmalar, fasonIsler, fasonHareketler }) {
  const [form, setForm] = useState({ firmaId: "", projeKodu: "", projeAdi: "", miktar: "", ucret: "", resimRef: "", aciklama: "" });
  const [msg, setMsg] = useState("");
  const [f, setF] = useState({ arama: "", firmaId: "", durum: "" });
  const [genisletilen, setGenisletilen] = useState(new Set());
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  const [iceMsg, setIceMsg] = useState("");
  const dosyaRef = useRef(null);
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const ekle = async () => {
    if (!form.firmaId || !form.projeAdi.trim()) { setMsg("Firma ve Proje Adı zorunlu."); setTimeout(() => setMsg(""), 2500); return; }
    await addDoc(collection(db, "fason_isler"), { ...form, durum: "bekliyor", olusturmaTarihi: todayISO() });
    setForm({ firmaId: form.firmaId, projeKodu: "", projeAdi: "", miktar: "", ucret: "", resimRef: "", aciklama: "" });
    setMsg("İş eklendi."); setTimeout(() => setMsg(""), 2000);
  };
  const sil = async (id) => {
    if (!window.confirm("Bu iş silinecek. Bağlı hareketler silinmez ama bağlantısız kalır. Emin misiniz?")) return;
    await deleteDoc(doc(db, "fason_isler", id));
  };
  const durumDegistir = async (id, durum) => { await updateDoc(doc(db, "fason_isler", id), { durum }); };
  const kaliteDegistir = async (id, mevcutKalite, kalite) => { await updateDoc(doc(db, "fason_isler", id), { kaliteDurumu: mevcutKalite === kalite ? "" : kalite }); };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true); setIceMsg("");
    try {
      const kayitlar = await excelDenFasonIsOku(dosya, fasonFirmalar);
      const gecerliler = kayitlar.filter((k) => k.firmaId);
      const atlanan = kayitlar.length - gecerliler.length;
      const { basarili, basarisiz } = await guvenliTopluYaz("fason_isler", gecerliler.map(({ firmaAdiGecici, ...rest }) => rest));
      setIceMsg(`${basarili} iş eklendi${atlanan > 0 ? `, ${atlanan} satır atlandı (firma bulunamadı, önce Firmalar'a ekleyin)` : ""}${basarisiz > 0 ? `, ${basarisiz} başarısız` : ""}.`);
    } catch (err) { console.error(err); setIceMsg("Hata: " + (err?.message || "bilinmeyen hata")); }
    setIceAktariliyor(false); e.target.value = ""; setTimeout(() => setIceMsg(""), 7000);
  };
  const disaAktar = () => excelIndir(
    fasonIsler.map((j) => ({
      "Firma Adı": fasonFirmalar.find((f) => f.id === j.firmaId)?.ad || "", "Proje Kodu": j.projeKodu, "Proje Adı": j.projeAdi,
      "Miktar": j.miktar, "Ücret": j.ucret, "Resim Referansı": j.resimRef, "Açıklama": j.aciklama,
      "Durum": FASON_DURUM[j.durum]?.label || "", "Oluşturma Tarihi": j.olusturmaTarihi,
    })), "fason-isler.xlsx", "İşler"
  );

  const filtrelenmis = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return fasonIsler.filter((j) => {
      if (f.firmaId && j.firmaId !== f.firmaId) return false;
      if (f.durum && j.durum !== f.durum) return false;
      if (q && !((j.projeKodu || "").toLowerCase().includes(q) || (j.projeAdi || "").toLowerCase().includes(q) || (j.resimRef || "").toLowerCase().includes(q))) return false;
      return true;
    });
  }, [fasonIsler, f]);

  const gruplar = useMemo(() => {
    const map = new Map(); const sira = [];
    filtrelenmis.forEach((j) => {
      const kod = (j.projeKodu || "").trim();
      const key = kod ? `${j.firmaId}::${kod.toLowerCase()}` : `tek::${j.id}`;
      if (!map.has(key)) { map.set(key, []); sira.push(key); }
      map.get(key).push(j);
    });
    return sira.map((key) => ({ key, isler: map.get(key) })).sort((a, b) => (b.isler[0]?.olusturmaTarihi || "").localeCompare(a.isler[0]?.olusturmaTarihi || ""));
  }, [filtrelenmis]);

  const grupToggle = (key) => setGenisletilen((s) => { const y = new Set(s); if (y.has(key)) y.delete(key); else y.add(key); return y; });
  const hammaddeGonderildiMi = (isId) => fasonHareketler.some((m) => m.isId === isId && m.tip === "giden");

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Yeni İş Ekle</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn-ghost" onClick={() => sablonIndir(["FİRMA ADI", "PROJE KODU", "PROJE ADI", "MİKTAR", "ÜCRET", "RESİM REFERANSI", "AÇIKLAMA", "DURUM", "OLUŞTURMA TARİHİ"], [["Örnek Fason Ltd.", "PRJ-001", "Örnek Proje", "100", "5000", "TR-001", "", "Bekliyor", "2026-01-15"]], "fason-is-sablonu.xlsx", "Şablon")}><FileDown size={14} /> Şablon İndir</button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}><Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excelden İçeri Al"}</button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excele Aktar</button>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div>
            <label className="field-label">Firma</label>
            <select className="input" value={form.firmaId} onChange={(e) => setForm((s) => ({ ...s, firmaId: e.target.value }))}>
              <option value="">Seçin…</option>
              {fasonFirmalar.map((f) => <option key={f.id} value={f.id}>{f.ad}{f.not && f.not.trim() ? ` — ${f.not}` : ""}</option>)}
            </select>
          </div>
          <div><label className="field-label">Proje Kodu</label><input className="input" value={form.projeKodu} onChange={(e) => setForm((s) => ({ ...s, projeKodu: e.target.value }))} /></div>
          <div><label className="field-label">Proje Adı</label><input className="input" value={form.projeAdi} onChange={(e) => setForm((s) => ({ ...s, projeAdi: e.target.value }))} /></div>
          <div><label className="field-label">Miktar</label><input className="input" value={form.miktar} onChange={(e) => setForm((s) => ({ ...s, miktar: e.target.value }))} /></div>
          <div><label className="field-label">Ücret (₺)</label><input className="input" type="number" step="0.01" value={form.ucret} onChange={(e) => setForm((s) => ({ ...s, ucret: e.target.value }))} /></div>
          <div><label className="field-label">Resim Referansı / Link</label><input className="input" value={form.resimRef} onChange={(e) => setForm((s) => ({ ...s, resimRef: e.target.value }))} /></div>
          <div style={{ gridColumn: "span 2" }}><label className="field-label">Açıklama</label><input className="input" value={form.aciklama} onChange={(e) => setForm((s) => ({ ...s, aciklama: e.target.value }))} /></div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button onClick={ekle} disabled={fasonFirmalar.length === 0} style={{ display: "flex", alignItems: "center", gap: 7, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}><Plus size={16} /> İş Ekle</button>
          {msg && <span style={{ fontSize: 12.5, color: "#8b929a" }}>{msg}</span>}
        </div>
        {fasonFirmalar.length === 0 && <div style={{ marginTop: 10, fontSize: 12, color: "#e8a33d" }}>Önce Firmalar sekmesinden firma ekleyin.</div>}
        {iceMsg && <div style={{ marginTop: 10, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{iceMsg}</div>}
      </div>

      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1", position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Proje kodu, adı, resim ref ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Firma</label>
            <select className="input" value={f.firmaId} onChange={setF2("firmaId")}>
              <option value="">Tümü</option>
              {fasonFirmalar.map((fm) => <option key={fm.id} value={fm.id}>{fm.ad}{fm.not && fm.not.trim() ? ` — ${fm.not}` : ""}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Durum</label>
            <select className="input" value={f.durum} onChange={setF2("durum")}>
              <option value="">Tümü</option>
              {Object.entries(FASON_DURUM).map(([k, d]) => <option key={k} value={k}>{d.label}</option>)}
            </select>
          </div>
        </div>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>İşler ({gruplar.reduce((s, g) => s + g.isler.length, 0)})</div>
        <div style={{ overflowX: "auto", maxHeight: 640, overflowY: "auto" }}>
          <table>
            <thead><tr><th></th><th>Proje / Kod</th><th>Firma</th><th>Miktar</th><th>Ücret</th><th>Durum</th><th>Hammadde</th><th>Kalite</th><th></th></tr></thead>
            <tbody>
              {gruplar.length === 0 && <tr><td colSpan={9} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>İş bulunamadı.</td></tr>}
              {gruplar.map((g) => {
                if (g.isler.length === 1) {
                  const j = g.isler[0];
                  const firma = fasonFirmalar.find((f) => f.id === j.firmaId);
                  const durum = FASON_DURUM[j.durum] || FASON_DURUM.bekliyor;
                  const gonderildi = hammaddeGonderildiMi(j.id);
                  return (
                    <tr key={j.id}>
                      <td></td>
                      <td>{j.projeKodu ? `${j.projeKodu} · ` : ""}{j.projeAdi}</td>
                      <td>{firma?.ad || "—"}</td>
                      <td style={{ fontFamily: "monospace" }}>{j.miktar || "—"}</td>
                      <td style={{ fontFamily: "monospace" }}>{j.ucret ? paraTR(j.ucret) : "—"}</td>
                      <td>
                        <select className="input" style={{ padding: "4px 6px", fontSize: 11.5 }} value={j.durum} onChange={(e) => durumDegistir(j.id, e.target.value)}>
                          {Object.entries(FASON_DURUM).map(([k, d]) => <option key={k} value={k}>{d.label}</option>)}
                        </select>
                      </td>
                      <td>{gonderildi ? <span className="pill">✓ Gönderildi</span> : <span className="pill" style={{ background: "#3a1f1f", color: "#e07a6b", borderColor: "#5a2a2a" }}>✕ Yok</span>}</td>
                      <td style={{ display: "flex", gap: 4 }}>
                        {Object.entries(FASON_KALITE).map(([k, q]) => (
                          <button key={k} onClick={() => kaliteDegistir(j.id, j.kaliteDurumu, k)} title={q.label} style={{ width: 20, height: 20, borderRadius: 5, border: `1px solid ${j.kaliteDurumu === k ? q.renk : "#2a4b52"}`, background: j.kaliteDurumu === k ? q.renk : "transparent", cursor: "pointer" }} />
                        ))}
                      </td>
                      <td><button onClick={() => sil(j.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                    </tr>
                  );
                }
                const acik = genisletilen.has(g.key);
                const firma = fasonFirmalar.find((f) => f.id === g.isler[0].firmaId);
                const sentCount = g.isler.filter((j) => hammaddeGonderildiMi(j.id)).length;
                return (
                  <React.Fragment key={g.key}>
                    <tr onClick={() => grupToggle(g.key)} style={{ cursor: "pointer", background: "#16232a" }}>
                      <td>{acik ? <ChevronDown size={14} /> : <ChevronRight size={14} />}</td>
                      <td colSpan={3}><span className="pill">{g.isler[0].projeKodu}</span> <span style={{ color: "#6b7178", fontSize: 12 }}>{g.isler.length} kalem</span></td>
                      <td colSpan={2} style={{ fontSize: 12 }}>{firma?.ad}</td>
                      <td colSpan={3} style={{ fontFamily: "monospace" }}>Hammadde: {sentCount} / {g.isler.length}</td>
                    </tr>
                    {acik && g.isler.map((j) => {
                      const durum = FASON_DURUM[j.durum] || FASON_DURUM.bekliyor;
                      const gonderildi = hammaddeGonderildiMi(j.id);
                      return (
                        <tr key={j.id}>
                          <td></td>
                          <td>{j.projeAdi}</td>
                          <td></td>
                          <td style={{ fontFamily: "monospace" }}>{j.miktar || "—"}</td>
                          <td style={{ fontFamily: "monospace" }}>{j.ucret ? paraTR(j.ucret) : "—"}</td>
                          <td>
                            <select className="input" style={{ padding: "4px 6px", fontSize: 11.5 }} value={j.durum} onChange={(e) => durumDegistir(j.id, e.target.value)}>
                              {Object.entries(FASON_DURUM).map(([k, d]) => <option key={k} value={k}>{d.label}</option>)}
                            </select>
                          </td>
                          <td>{gonderildi ? <span className="pill">✓</span> : <span className="pill" style={{ background: "#3a1f1f", color: "#e07a6b", borderColor: "#5a2a2a" }}>✕</span>}</td>
                          <td style={{ display: "flex", gap: 4 }}>
                            {Object.entries(FASON_KALITE).map(([k, q]) => (
                              <button key={k} onClick={() => kaliteDegistir(j.id, j.kaliteDurumu, k)} title={q.label} style={{ width: 20, height: 20, borderRadius: 5, border: `1px solid ${j.kaliteDurumu === k ? q.renk : "#2a4b52"}`, background: j.kaliteDurumu === k ? q.renk : "transparent", cursor: "pointer" }} />
                            ))}
                          </td>
                          <td><button onClick={() => sil(j.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                        </tr>
                      );
                    })}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Fason Hareketler ----------
function FasonHareketler({ fasonFirmalar, fasonIsler, fasonHareketler }) {
  const [form, setForm] = useState({ isId: "", tip: "giden", urunAdi: "", malzemeCinsi: "", kalite: "", aciklama: "", miktar: "", birim: "", birimFiyat: "", tarih: todayISO(), not: "" });
  const [msg, setMsg] = useState("");
  const [f, setF] = useState({ arama: "", tip: "", firmaId: "" });
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  const [iceMsg, setIceMsg] = useState("");
  const dosyaRef = useRef(null);
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const isLabel = (isId) => {
    const j = fasonIsler.find((x) => x.id === isId);
    if (!j) return "—";
    const firma = fasonFirmalar.find((f) => f.id === j.firmaId);
    return `${j.projeKodu ? j.projeKodu + " · " : ""}${j.projeAdi} (${firma?.ad || "?"})`;
  };

  const ekle = async () => {
    if (!form.isId || !form.urunAdi.trim() || !form.miktar) { setMsg("İş, ürün/malzeme adı ve miktar zorunlu."); setTimeout(() => setMsg(""), 2500); return; }
    await addDoc(collection(db, "fason_hareketler"), { ...form, miktar: Number(form.miktar) || 0, birimFiyat: Number(form.birimFiyat) || 0 });
    setForm({ isId: form.isId, tip: form.tip, urunAdi: "", malzemeCinsi: "", kalite: "", aciklama: "", miktar: "", birim: "", birimFiyat: "", tarih: form.tarih, not: "" });
    setMsg("Hareket eklendi."); setTimeout(() => setMsg(""), 2000);
  };
  const sil = async (id) => { await deleteDoc(doc(db, "fason_hareketler", id)); };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true); setIceMsg("");
    try {
      const kayitlar = await excelDenFasonHareketOku(dosya, fasonFirmalar, fasonIsler);
      const { basarili, basarisiz } = await guvenliTopluYaz("fason_hareketler", kayitlar);
      setIceMsg(`${basarili} hareket eklendi${basarisiz > 0 ? `, ${basarisiz} başarısız` : ""}.`);
    } catch (err) { console.error(err); setIceMsg("Hata: " + (err?.message || "bilinmeyen hata")); }
    setIceAktariliyor(false); e.target.value = ""; setTimeout(() => setIceMsg(""), 7000);
  };
  const disaAktar = () => excelIndir(
    fasonHareketler.map((m) => {
      const j = fasonIsler.find((x) => x.id === m.isId);
      const firma = j ? fasonFirmalar.find((f) => f.id === j.firmaId) : null;
      return {
        "Firma Adı": firma?.ad || "", "Proje Kodu": j?.projeKodu || "", "Proje Adı": j?.projeAdi || "",
        "Tip": m.tip === "giden" ? "Giden (Hammadde)" : "Gelen (Ürün/Fason)", "Ürün / Malzeme Adı": m.urunAdi,
        "Malzeme Cinsi": m.malzemeCinsi, "Kalite": m.kalite, "Açıklama": m.aciklama,
        "Miktar": m.miktar, "Birim": m.birim, "Birim Fiyat": m.birimFiyat,
        "Tutar": (Number(m.miktar) || 0) * (Number(m.birimFiyat) || 0), "Tarih": m.tarih, "Not": m.not,
      };
    }), "fason-hareketler.xlsx", "Hareketler"
  );

  const filtrelenmis = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return [...fasonHareketler].filter((m) => {
      const j = fasonIsler.find((x) => x.id === m.isId);
      if (f.tip && m.tip !== f.tip) return false;
      if (f.firmaId && (!j || j.firmaId !== f.firmaId)) return false;
      if (q && !((m.urunAdi || "").toLowerCase().includes(q) || (m.malzemeCinsi || "").toLowerCase().includes(q) || (m.kalite || "").toLowerCase().includes(q) || (j?.projeKodu || "").toLowerCase().includes(q))) return false;
      return true;
    }).sort((a, b) => (b.tarih || "").localeCompare(a.tarih || ""));
  }, [fasonHareketler, fasonIsler, f]);

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Yeni Hareket Ekle</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn-ghost" onClick={() => sablonIndir(["FİRMA ADI", "PROJE KODU", "PROJE ADI", "TİP", "ÜRÜN/MALZEME ADI", "MALZEME CİNSİ", "KALİTE", "AÇIKLAMA", "MİKTAR", "BİRİM", "BİRİM FİYAT", "TARİH", "NOT"], [["Örnek Fason Ltd.", "PRJ-001", "Örnek Proje", "Giden (Hammadde)", "Çelik Sac", "Paslanmaz Çelik", "304", "", "50", "kg", "120", "2026-01-15", ""]], "fason-hareket-sablonu.xlsx", "Şablon")}><FileDown size={14} /> Şablon İndir</button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}><Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excelden İçeri Al"}</button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excele Aktar</button>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "span 2" }}>
            <label className="field-label">İş</label>
            <select className="input" value={form.isId} onChange={(e) => setForm((s) => ({ ...s, isId: e.target.value }))}>
              <option value="">Seçin…</option>
              {fasonIsler.map((j) => <option key={j.id} value={j.id}>{isLabel(j.id)}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Hareket Tipi</label>
            <select className="input" value={form.tip} onChange={(e) => setForm((s) => ({ ...s, tip: e.target.value }))}>
              <option value="giden">Giden — Hammadde (firmaya)</option>
              <option value="gelen">Gelen — Ürün / Fason (firmadan)</option>
            </select>
          </div>
          <div><label className="field-label">{form.tip === "giden" ? "Malzeme İsmi" : "Ürün Adı"}</label><input className="input" value={form.urunAdi} onChange={(e) => setForm((s) => ({ ...s, urunAdi: e.target.value }))} /></div>
          {form.tip === "giden" && (
            <>
              <div><label className="field-label">Malzeme Cinsi</label><input className="input" value={form.malzemeCinsi} onChange={(e) => setForm((s) => ({ ...s, malzemeCinsi: e.target.value }))} /></div>
              <div><label className="field-label">Kalite</label><input className="input" value={form.kalite} onChange={(e) => setForm((s) => ({ ...s, kalite: e.target.value }))} /></div>
            </>
          )}
          <div><label className="field-label">Miktar</label><input className="input" type="number" step="0.01" value={form.miktar} onChange={(e) => setForm((s) => ({ ...s, miktar: e.target.value }))} /></div>
          <div><label className="field-label">Birim</label><input className="input" value={form.birim} onChange={(e) => setForm((s) => ({ ...s, birim: e.target.value }))} placeholder="kg, adet, mt" /></div>
          <div><label className="field-label">Birim Fiyat (₺)</label><input className="input" type="number" step="0.01" value={form.birimFiyat} onChange={(e) => setForm((s) => ({ ...s, birimFiyat: e.target.value }))} /></div>
          <div><label className="field-label">Tarih</label><input className="input" type="date" value={form.tarih} onChange={(e) => setForm((s) => ({ ...s, tarih: e.target.value }))} /></div>
          <div style={{ gridColumn: "span 2" }}><label className="field-label">Not</label><input className="input" value={form.not} onChange={(e) => setForm((s) => ({ ...s, not: e.target.value }))} /></div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button onClick={ekle} disabled={fasonIsler.length === 0} style={{ display: "flex", alignItems: "center", gap: 7, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}><Plus size={16} /> Hareketi Kaydet</button>
          {msg && <span style={{ fontSize: 12.5, color: "#8b929a" }}>{msg}</span>}
        </div>
        {fasonIsler.length === 0 && <div style={{ marginTop: 10, fontSize: 12, color: "#e8a33d" }}>Önce İşler sekmesinden bir iş ekleyin.</div>}
        {iceMsg && <div style={{ marginTop: 10, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{iceMsg}</div>}
      </div>

      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1", position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Malzeme/ürün, cins, kalite, proje kodu ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Tip</label>
            <select className="input" value={f.tip} onChange={setF2("tip")}>
              <option value="">Tümü</option><option value="giden">Giden</option><option value="gelen">Gelen</option>
            </select>
          </div>
          <div>
            <label className="field-label">Firma</label>
            <select className="input" value={f.firmaId} onChange={setF2("firmaId")}>
              <option value="">Tümü</option>
              {fasonFirmalar.map((fm) => <option key={fm.id} value={fm.id}>{fm.ad}{fm.not && fm.not.trim() ? ` — ${fm.not}` : ""}</option>)}
            </select>
          </div>
        </div>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Hareketler ({filtrelenmis.length})</div>
        <div style={{ overflowX: "auto", maxHeight: 640, overflowY: "auto" }}>
          <table>
            <thead><tr><th>Tarih</th><th>Tip</th><th>Ürün / Malzeme</th><th>İş</th><th>Miktar</th><th>Tutar</th><th></th></tr></thead>
            <tbody>
              {filtrelenmis.length === 0 && <tr><td colSpan={7} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Hareket bulunamadı.</td></tr>}
              {filtrelenmis.map((m) => {
                const tutar = (Number(m.miktar) || 0) * (Number(m.birimFiyat) || 0);
                return (
                  <tr key={m.id}>
                    <td style={{ fontFamily: "monospace" }}>{m.tarih}</td>
                    <td>{m.tip === "giden" ? <span className="pill">↑ Giden</span> : <span className="pill" style={{ background: "#113330", color: "#4b8f5e", borderColor: "#1f4d47" }}>↓ Gelen</span>}</td>
                    <td>{m.urunAdi}{m.kalite ? ` · ${m.kalite}` : ""}</td>
                    <td style={{ fontSize: 12 }}>{isLabel(m.isId)}</td>
                    <td style={{ fontFamily: "monospace" }}>{m.miktar} {m.birim}</td>
                    <td style={{ fontFamily: "monospace", fontWeight: 700, color: m.tip === "giden" ? "#e8a33d" : "#4b8f5e" }}>{paraTR(tutar)}</td>
                    <td><button onClick={() => sil(m.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Fason Hatırlatıcılar ----------
function FasonHatirlaticilar({ fasonIsler, fasonHatirlaticilar }) {
  const [form, setForm] = useState({ baslik: "", tarih: todayISO(), isId: "", not: "" });
  const [msg, setMsg] = useState("");
  const [filtre, setFiltre] = useState("bekleyen");
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  const [iceMsg, setIceMsg] = useState("");
  const dosyaRef = useRef(null);

  const ekle = async () => {
    if (!form.baslik.trim()) { setMsg("Başlık zorunlu."); setTimeout(() => setMsg(""), 2000); return; }
    await addDoc(collection(db, "fason_hatirlaticilar"), { ...form, isId: form.isId || null, tamamlandi: false });
    setForm({ baslik: "", tarih: todayISO(), isId: "", not: "" });
    setMsg("Hatırlatıcı eklendi."); setTimeout(() => setMsg(""), 2000);
  };
  const sil = async (id) => { await deleteDoc(doc(db, "fason_hatirlaticilar", id)); };
  const toggle = async (r) => { await updateDoc(doc(db, "fason_hatirlaticilar", r.id), { tamamlandi: !r.tamamlandi }); };
  const disaAktar = () => excelIndir(
    fasonHatirlaticilar.map((r) => ({ "Başlık": r.baslik, "Tarih": r.tarih, "Durum": r.tamamlandi ? "Tamamlandı" : "Bekliyor", "Not": r.not })),
    "fason-hatirlaticilar.xlsx", "Hatırlatıcılar"
  );

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true); setIceMsg("");
    try {
      const rows = await dosyaOku(dosya);
      let baslangic = 0;
      const ilkSatir = (rows[0] || []).map((v) => String(v || "").trim().toLowerCase());
      if (ilkSatir[0] && ilkSatir[0].includes("başlık")) baslangic = 1;
      const kayitlar = [];
      for (let i = baslangic; i < rows.length; i++) {
        const r = rows[i] || [];
        const baslik = String(r[0] || "").trim();
        if (!baslik) continue;
        const projeAdi = String(r[2] || "").trim();
        const is = projeAdi ? fasonIsler.find((j) => j.projeAdi.toLowerCase() === projeAdi.toLowerCase()) : null;
        kayitlar.push({
          baslik, tarih: String(r[1] || "").trim(),
          isId: is ? is.id : null,
          not: String(r[3] || "").trim(),
          tamamlandi: String(r[4] || "").trim().toLowerCase().includes("tamam"),
        });
      }
      const { basarili, basarisiz } = await guvenliTopluYaz("fason_hatirlaticilar", kayitlar);
      setIceMsg(`${basarili} hatırlatıcı eklendi${basarisiz > 0 ? `, ${basarisiz} başarısız` : ""}.`);
    } catch (err) {
      console.error(err);
      setIceMsg("Hata: " + (err?.message || "bilinmeyen hata"));
    }
    setIceAktariliyor(false); e.target.value = ""; setTimeout(() => setIceMsg(""), 6000);
  };

  const filtrelenmis = useMemo(() => {
    let liste = [...fasonHatirlaticilar];
    if (filtre === "bekleyen") liste = liste.filter((r) => !r.tamamlandi);
    else if (filtre === "tamamlanan") liste = liste.filter((r) => r.tamamlandi);
    return liste.sort((a, b) => (a.tarih || "").localeCompare(b.tarih || ""));
  }, [fasonHatirlaticilar, filtre]);

  const bugun = todayISO();

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Yeni Hatırlatıcı</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn-ghost" onClick={() => sablonIndir(["Başlık", "Tarih", "İlişkili Proje Adı", "Not", "Durum"], [["Teslimat kontrolü", "2026-01-20", "Örnek Proje", "", "Bekliyor"]], "fason-hatirlatici-sablonu.xlsx", "Şablon")}><FileDown size={14} /> Şablon İndir</button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}><Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excel'den İçe Aktar"}</button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excele Aktar</button>
          </div>
        </div>
        {iceMsg && <div style={{ marginBottom: 14, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{iceMsg}</div>}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div><label className="field-label">Başlık</label><input className="input" value={form.baslik} onChange={(e) => setForm((s) => ({ ...s, baslik: e.target.value }))} /></div>
          <div><label className="field-label">Tarih</label><input className="input" type="date" value={form.tarih} onChange={(e) => setForm((s) => ({ ...s, tarih: e.target.value }))} /></div>
          <div>
            <label className="field-label">İşle İlişkilendir (opsiyonel)</label>
            <select className="input" value={form.isId} onChange={(e) => setForm((s) => ({ ...s, isId: e.target.value }))}>
              <option value="">Yok</option>
              {fasonIsler.map((j) => <option key={j.id} value={j.id}>{j.projeAdi}</option>)}
            </select>
          </div>
          <div style={{ gridColumn: "span 2" }}><label className="field-label">Not</label><input className="input" value={form.not} onChange={(e) => setForm((s) => ({ ...s, not: e.target.value }))} /></div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button onClick={ekle} style={{ display: "flex", alignItems: "center", gap: 7, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}><Plus size={16} /> Ekle</button>
          {msg && <span style={{ fontSize: 12.5, color: "#8b929a" }}>{msg}</span>}
        </div>
      </div>

      <div style={{ display: "flex", gap: 8 }}>
        {[{ id: "bekleyen", label: "Bekleyen" }, { id: "tamamlanan", label: "Tamamlanan" }, { id: "tumu", label: "Tümü" }].map((s) => (
          <button key={s.id} onClick={() => setFiltre(s.id)} style={{ flex: 1, padding: "9px 14px", borderRadius: 8, cursor: "pointer", fontWeight: 700, fontSize: 12.5, background: filtre === s.id ? "#2dd4bf" : "#1b333c", color: filtre === s.id ? "#142a30" : "#c7cbd1", border: `1px solid ${filtre === s.id ? "#2dd4bf" : "#2a4b52"}` }}>{s.label}</button>
        ))}
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        {filtrelenmis.length === 0 ? (
          <div style={{ color: "#6b7178", textAlign: "center", padding: 32, fontSize: 13.5 }}>Hatırlatıcı bulunamadı.</div>
        ) : (
          filtrelenmis.map((r) => {
            const gecikti = !r.tamamlandi && r.tarih && r.tarih < bugun;
            const bugunMu = !r.tamamlandi && r.tarih === bugun;
            const is = r.isId ? fasonIsler.find((j) => j.id === r.isId) : null;
            return (
              <div key={r.id} style={{ padding: "14px 20px", borderBottom: "1px solid #223b42", display: "flex", alignItems: "flex-start", gap: 12, opacity: r.tamamlandi ? 0.6 : 1 }}>
                <button onClick={() => toggle(r)} style={{ width: 20, height: 20, borderRadius: "999px", border: `1.5px solid ${r.tamamlandi ? "#4b8f5e" : "#2a4b52"}`, background: r.tamamlandi ? "#4b8f5e" : "transparent", color: "#fff", cursor: "pointer", flexShrink: 0, fontSize: 12 }}>{r.tamamlandi ? "✓" : ""}</button>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13.5, textDecoration: r.tamamlandi ? "line-through" : "none" }}>{r.baslik}</div>
                  <div style={{ fontSize: 11.5, color: gecikti ? "#e07a6b" : bugunMu ? "#e8a33d" : "#6b7178", marginTop: 2 }}>{r.tarih}{gecikti ? " · gecikti" : bugunMu ? " · bugün" : ""}</div>
                  {is && <div style={{ fontSize: 11.5, color: "#6b7178", marginTop: 2 }}>{is.projeAdi}</div>}
                  {r.not && <div style={{ fontSize: 12, color: "#8b929a", marginTop: 2 }}>{r.not}</div>}
                </div>
                <button onClick={() => sil(r.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button>
              </div>
            );
          })
        )}
      </div>
    </div>
  );
}

// ---------- Yardım ----------
const YARDIM_ICERIK = [
  {
    grup: "Ana Sayfa", tab: "ana-sayfa",
    ogeler: [
      { baslik: "Ana Sayfa nedir?", detay: "Giriş yapınca karşına çıkan kontrol paneli. Üstte gün/tarih ve isim düzenleme, altında biten depo stoğu / geciken hatırlatıcı gibi uyarılar, sonra her modülün özet kartları (tıklayınca o ekrana götürür), en altta genel istatistikler var." },
    ],
  },
  {
    grup: "Hammadde", tab: "hammadde-kayit",
    ogeler: [
      { baslik: "Hammadde Kaydı — sipariş girme", detay: "Cari (tedarikçi) isim, proje kodu/adı, kalite, ölçü/adet açıklamaları ve miktar (kg) girip \"Açık Sipariş Olarak Ekle\" dersin. Kayıt \"Açık Siparişler\" listesine düşer. Sipariş gelince \"Tamamlandı\" işaretleyip \"Tamamlanan\" sekmesine taşıyabilirsin. Excel şablonuyla toplu da yükleyebilirsin." },
      { baslik: "Hammadde Raporu", detay: "Açık ve tamamlanan siparişlerin toplu görünümü, cari/proje bazında filtrelenebilir, Excel'e aktarılabilir." },
      { baslik: "Hammadde Kayıtları Sil", detay: "Yanlış girilmiş hammadde kayıtlarını seçip toplu silme ekranı." },
    ],
  },
  {
    grup: "Metal Ölçü", tab: "metal-hizli",
    ogeler: [
      { baslik: "Hızlı KG Hesabı", detay: "Malzeme (Çelik, Paslanmaz, Bronz, Kestamid, Alüminyum — otomatik hazır gelir) ve kesit türü (Mil/Kare/Lama/Boru) seçip ölçüleri, boyu ve adedi girersin; parça ve toplam ağırlık anında hesaplanır. \"Kaydet\" dersen \"Geçmiş Ölçümler\"e düşer. Excel'den toplu ölçüm de yükleyebilirsin." },
      { baslik: "Geçmiş Ölçümler", detay: "Kaydettiğin tüm ölçümlerin listesi — malzeme/tür/tarihe göre filtrelenebilir, Excel'e aktarılabilir, tek tek silinebilir." },
      { baslik: "Metal Ölçü Raporu", detay: "Malzeme bazında toplam kg/tutar dağılımı ve talep (proje) bazında gruplu detaylı rapor." },
      { baslik: "Malzeme Tanımları", detay: "Varsayılan malzemelere ek olarak kendi özel malzemeni (adı + yoğunluğu) tanımlayabileceğin ekran. Genelde gerek kalmaz, varsayılanlar zaten hazır gelir." },
    ],
  },
  {
    grup: "Depo Stok", tab: "depo-kart",
    ogeler: [
      { baslik: "Stok Kartı Oluştur", detay: "Yeni bir malzeme/parçayı sisteme ilk kez tanımladığın yer: Stok Kodu, Stok İsmi, Birim, Ana Grup, Alt Grup girilir. Excel'den toplu yükleme yapabilirsin (başlık isimleri eşleştiği sürece sütun sırası önemli değil). Listede miktar hücresine tıklayıp birden fazla satırı düzenleyip \"Değişiklikleri Kaydet\" ile toplu kaydedebilirsin. Sarı kutudaki \"Toplu Miktar Ayarla\" ile TÜM kalemleri tek seferde aynı sayıya çekebilirsin." },
      { baslik: "Depo Giriş", detay: "Depoya gelen malzemeyi var olan bir stok kartına ekler (miktarı artırır). Alt Grup filtresi ya da arama ile stoğu bul, listeden seç, miktar/tarih/açıklama gir, \"Girişi Kaydet\" de. Aynı ekranda \"Toplu Stok Girişi\" tablosuyla birden fazla kalemin miktarını aynı anda girip tek seferde kaydedebilirsin." },
      { baslik: "Depo Çıkış", detay: "Depodan verilen malzemeyi kaydeder, stoktan otomatik düşer. Stok seç, hangi makineye gittiğini seç (zorunlu), miktar gir, \"Çıkışı Kaydet\" de. Excel'den toplu çıkış da yapılabilir." },
      { baslik: "Stok Hareketleri", detay: "Tüm giriş/çıkış geçmişi — Bugün/Bu Ay/Bu Yıl hızlı filtreleriyle, örneğin günlük çıkışları çekip Mikro'ya elle işlemek için Excel'e aktarabilirsin. Hedef Makine sütunu da burada görünür." },
      { baslik: "Depo Stok Raporu", detay: "Güncel stok durumu, biten/eksi stok uyarısı, hangi makineye ne kadar çıkış yapıldığı (Makine Bazında Çıkış tablosu), kalem bazında hareket özeti — hepsi dönem filtreli ve Excel'e aktarılabilir." },
      { baslik: "Stok Kartı Sil", detay: "Yanlış/gereksiz stok kartlarını seçip toplu silme ekranı. Hareket geçmişini silmez, sadece kartı kaldırır." },
    ],
  },
  {
    grup: "Fason Takip", tab: "fason-ozet",
    ogeler: [
      { baslik: "Özet", detay: "Fason (dış üretim) firmalarının bakiye durumu, toplam giden hammadde / gelen ürün tutarları, yaklaşan hatırlatıcılar." },
      { baslik: "Firmalar", detay: "Fason çalıştığın firmaları (isim, yetkili, not) tanımladığın ve listelediğin yer." },
      { baslik: "İşler", detay: "Bir firmaya verdiğin işi (proje kodu/adı, miktar, ücret, durum: Bekliyor/Üretimde/Tamamlandı) kaydettiğin yer. Aynı proje koduna sahip işler otomatik gruplanır. Kalite kontrolü (Okeylendi/Red/Ölçümde) de buradan işaretlenir." },
      { baslik: "Hareketler", detay: "Bir işe bağlı olarak giden (hammadde) ya da gelen (ürün/fason bedeli) hareketleri kaydettiğin yer, tutar otomatik hesaplanır." },
      { baslik: "Hatırlatıcılar", detay: "Bir işle ilişkilendirilebilen tarihli hatırlatıcılar — geciken/bugünkü olanlar renkli işaretlenir, Ana Sayfa'da da uyarı olarak çıkar." },
      { baslik: "Fason Takip Raporu", detay: "Toplam giden/gelen/net bakiye, iş durum dağılımı, kalite kontrolü dağılımı, firma bazında detaylı tablo." },
    ],
  },
  {
    grup: "Diğer", tab: "takimlar",
    ogeler: [
      { baslik: "Takımlar", detay: "Üretimde çalışan takım isimlerini tanımladığın basit liste ekranı." },
      { baslik: "Makineler", detay: "İşletmedeki makine isimlerini tanımladığın liste — bu liste, Depo Çıkış'taki \"Hedef Makine\" seçiminde kullanılır. Yeni bir makine varsa önce buraya eklemelisin." },
      { baslik: "Kullanıcılar", detay: "Ekip arkadaşlarına erişim vermek için: e-posta/şifre ile hesap oluşturabilir ya da Google hesabına giriş izni verebilirsin. Herkes kendi hesabıyla \"Kayıt Ol\" linkinden de kayıt olabilir." },
    ],
  },
];

function YardimEkrani({ git }) {
  const [arama, setArama] = useState("");
  const [acikGrup, setAcikGrup] = useState(null);

  const filtrelenmis = useMemo(() => {
    if (!arama.trim()) return YARDIM_ICERIK;
    const q = arama.trim().toLowerCase();
    return YARDIM_ICERIK.map((g) => ({
      ...g,
      ogeler: g.ogeler.filter((o) => o.baslik.toLowerCase().includes(q) || o.detay.toLowerCase().includes(q) || g.grup.toLowerCase().includes(q)),
    })).filter((g) => g.ogeler.length > 0);
  }, [arama]);

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 24, background: "linear-gradient(135deg, #1b333c 0%, #16232a 100%)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
          <HelpCircle size={20} color="#2dd4bf" />
          <div style={{ fontSize: 18, fontWeight: 800 }}>Yardım</div>
        </div>
        <div style={{ fontSize: 13, color: "#8b929a", lineHeight: 1.6 }}>
          SAKLAZ · ÜRETİM ERP'yi ilk kez kullanıyorsan aşağıdan modül modül ne işe yaradığını okuyabilirsin. Bir konuyu ara ya da başlığa tıklayıp aç, "Bu Ekrana Git" ile doğrudan o sayfaya geç.
        </div>
      </div>

      <div className="card" style={{ padding: "12px 16px" }}>
        <div style={{ position: "relative" }}>
          <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
          <input className="input" style={{ paddingLeft: 30 }} placeholder="Ne yapmak istiyorsun? Örn: makine ekleme, çıkış yapma, rapor…" value={arama} onChange={(e) => setArama(e.target.value)} />
        </div>
      </div>

      {filtrelenmis.length === 0 && (
        <div className="card" style={{ padding: 24, textAlign: "center", color: "#6b7178", fontSize: 13 }}>Aramanla eşleşen bir konu bulunamadı.</div>
      )}

      {filtrelenmis.map((g) => (
        <div key={g.grup} className="card" style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14, color: "#2dd4bf" }}>{g.grup}</div>
          {g.ogeler.map((o, i) => {
            const key = g.grup + i;
            const acik = acikGrup === key;
            return (
              <div key={key} style={{ borderBottom: "1px solid #223b42" }}>
                <button
                  onClick={() => setAcikGrup(acik ? null : key)}
                  style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 20px", background: "none", border: "none", cursor: "pointer", textAlign: "left" }}
                >
                  <span style={{ fontSize: 13.5, fontWeight: 500 }}>{o.baslik}</span>
                  {acik ? <ChevronDown size={16} color="#6b7178" /> : <ChevronRight size={16} color="#6b7178" />}
                </button>
                {acik && (
                  <div style={{ padding: "0 20px 16px" }}>
                    <div style={{ fontSize: 13, color: "#8b929a", lineHeight: 1.7, marginBottom: 10 }}>{o.detay}</div>
                    <button onClick={() => git(g.tab)} className="btn-ghost" style={{ fontSize: 12 }}>Bu Ekrana Git →</button>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      ))}

      <div className="card" style={{ padding: 18, fontSize: 12, color: "#6b7178" }}>
        Aradığını bulamadın mı ya da bir sorun mu var? Programı geliştiren ekibe (Claude ile) ulaşıp anlat, hemen düzeltilir ya da eklenir.
      </div>
    </div>
  );
}

// ---------- Kullanıcı Yönetimi ----------
function KullaniciYonetimi({ mevcutKullanici }) {
  const [kullanicilar, setKullanicilar] = useState([]);
  const [form, setForm] = useState({ ad: "", email: "", sifre: "" });
  const [googleForm, setGoogleForm] = useState({ ad: "", email: "" });
  const [olusturuluyor, setOlusturuluyor] = useState(false);
  const [googleEkleniyor, setGoogleEkleniyor] = useState(false);
  const [msg, setMsg] = useState("");
  const [msgTip, setMsgTip] = useState("bilgi"); // "bilgi" | "hata"
  const [sifirlamaGonderilen, setSifirlamaGonderilen] = useState("");
  const set = (k) => (e) => setForm((s) => ({ ...s, [k]: e.target.value }));
  const setG = (k) => (e) => setGoogleForm((s) => ({ ...s, [k]: e.target.value }));

  useEffect(() => {
    const unsub = onSnapshot(collection(db, "kullanicilar"), (snap) =>
      setKullanicilar(snap.docs.map((d) => ({ id: d.id, ...d.data() })).sort((a, b) => (b.eklenmeTarihi || 0) - (a.eklenmeTarihi || 0)))
    );
    return unsub;
  }, []);

  const kullaniciOlustur = async () => {
    if (!form.email || !form.sifre) {
      setMsgTip("hata"); setMsg("E-posta ve şifre zorunlu.");
      return;
    }
    if (form.sifre.length < 6) {
      setMsgTip("hata"); setMsg("Şifre en az 6 karakter olmalı (Firebase kuralı).");
      return;
    }
    setOlusturuluyor(true);
    setMsg("");
    try {
      await digerKullaniciOlustur(form.email.trim(), form.sifre);
      await addDoc(collection(db, "kullanicilar"), {
        ad: form.ad.trim(), email: form.email.trim(), emailKucuk: form.email.trim().toLowerCase(),
        tur: "sifreli", eklenmeTarihi: Date.now(),
      });
      setForm({ ad: "", email: "", sifre: "" });
      setMsgTip("bilgi"); setMsg(`${form.email} için hesap oluşturuldu. Bu bilgileri kişiye ilet.`);
    } catch (err) {
      const kod = err?.code || "";
      let aciklama = kod;
      if (kod.includes("email-already-in-use")) aciklama = "Bu e-posta zaten kayıtlı.";
      else if (kod.includes("invalid-email")) aciklama = "E-posta adresi geçersiz.";
      else if (kod.includes("weak-password")) aciklama = "Şifre çok zayıf, en az 6 karakter olmalı.";
      setMsgTip("hata"); setMsg("Hesap oluşturulamadı: " + aciklama);
    }
    setOlusturuluyor(false);
    setTimeout(() => setMsg(""), 6000);
  };

  const googleIzniEkle = async () => {
    if (!googleForm.email) {
      setMsgTip("hata"); setMsg("E-posta zorunlu.");
      return;
    }
    setGoogleEkleniyor(true);
    setMsg("");
    const emailKucuk = googleForm.email.trim().toLowerCase();
    const zatenVar = kullanicilar.some((k) => (k.emailKucuk || k.email?.toLowerCase()) === emailKucuk);
    if (zatenVar) {
      setMsgTip("hata"); setMsg("Bu e-posta zaten listede.");
      setGoogleEkleniyor(false);
      return;
    }
    try {
      await addDoc(collection(db, "kullanicilar"), {
        ad: googleForm.ad.trim(), email: googleForm.email.trim(), emailKucuk,
        tur: "google", eklenmeTarihi: Date.now(),
      });
      setGoogleForm({ ad: "", email: "" });
      setMsgTip("bilgi"); setMsg(`${googleForm.email} artık Google ile giriş yapabilir.`);
    } catch (err) {
      setMsgTip("hata"); setMsg("Eklenemedi: " + (err?.message || "bilinmeyen hata"));
    }
    setGoogleEkleniyor(false);
    setTimeout(() => setMsg(""), 6000);
  };

  const sifreSifirlaGonder = async (email) => {
    try {
      await sendPasswordResetEmail(auth, email);
      setSifirlamaGonderilen(email);
      setTimeout(() => setSifirlamaGonderilen(""), 4000);
    } catch (err) {
      setMsgTip("hata"); setMsg("Gönderilemedi: " + (err?.code || "bilinmeyen hata"));
      setTimeout(() => setMsg(""), 4000);
    }
  };

  const listedenKaldir = async (id) => { await deleteDoc(doc(db, "kullanicilar", id)); };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>Yeni Kullanıcı Oluştur (E-posta / Şifre)</div>
        <div style={{ fontSize: 12, color: "#6b7178", marginBottom: 16 }}>
          Kişinin kendi e-posta ve şifresiyle giriş yapabilmesi için hesap açar. Şifreyi kendine ya da kişiye iletmen gerekir.
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div><label className="field-label">Ad Soyad (opsiyonel)</label><input className="input" placeholder="Örn: Ahmet Yılmaz" value={form.ad} onChange={set("ad")} /></div>
          <div><label className="field-label">E-posta</label><input className="input" type="email" placeholder="ornek@firma.com" value={form.email} onChange={set("email")} /></div>
          <div><label className="field-label">Şifre</label><input className="input" type="text" placeholder="En az 6 karakter" value={form.sifre} onChange={set("sifre")} /></div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button onClick={kullaniciOlustur} disabled={olusturuluyor} style={{ display: "flex", alignItems: "center", gap: 7, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
            <UserPlus size={16} /> {olusturuluyor ? "Oluşturuluyor…" : "Kullanıcı Oluştur"}
          </button>
        </div>
      </div>

      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>Google ile Giriş İzni Ver</div>
        <div style={{ fontSize: 12, color: "#6b7178", marginBottom: 16 }}>
          Şifre oluşturmaz, sadece bu e-postanın kendi Google hesabıyla giriş yapmasına izin verir. Kişinin kullandığı Google hesabının e-postasını gir.
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div><label className="field-label">Ad Soyad (opsiyonel)</label><input className="input" placeholder="Örn: Ahmet Yılmaz" value={googleForm.ad} onChange={setG("ad")} /></div>
          <div><label className="field-label">Google E-postası</label><input className="input" type="email" placeholder="ornek@gmail.com" value={googleForm.email} onChange={setG("email")} /></div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button onClick={googleIzniEkle} disabled={googleEkleniyor} style={{ display: "flex", alignItems: "center", gap: 7, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
            <Chrome size={16} /> {googleEkleniyor ? "Ekleniyor…" : "İzin Ver"}
          </button>
        </div>
        {msg && (
          <div style={{ marginTop: 14, fontSize: 12.5, color: msgTip === "hata" ? "#e07a6b" : "#2dd4bf", background: msgTip === "hata" ? "#3a1f1f" : "#113330", border: `1px solid ${msgTip === "hata" ? "#5a2a2a" : "#1f4d47"}`, borderRadius: 7, padding: "9px 12px" }}>{msg}</div>
        )}
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Oluşturulan Kullanıcılar ({kullanicilar.length})</div>
        {kullanicilar.length === 0 ? (
          <div style={{ color: "#6b7178", textAlign: "center", padding: 32, fontSize: 13.5 }}>Henüz kullanıcı oluşturulmadı.</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table>
              <thead><tr><th>Ad Soyad</th><th>E-posta</th><th>Tür</th><th>Eklenme Tarihi</th><th></th><th></th></tr></thead>
              <tbody>
                {kullanicilar.map((k) => (
                  <tr key={k.id}>
                    <td>{k.ad || "—"}</td>
                    <td>{k.email}</td>
                    <td>
                      {k.tur === "google"
                        ? <span className="pill" style={{ background: "#1f2d3a", color: "#7fb0e0", borderColor: "#2c4a63" }}>Google</span>
                        : <span className="pill">Şifreli</span>}
                    </td>
                    <td style={{ fontFamily: "monospace" }}>{k.eklenmeTarihi ? new Date(k.eklenmeTarihi).toLocaleDateString("tr-TR") : "—"}</td>
                    <td>
                      {k.tur !== "google" && (
                        <button onClick={() => sifreSifirlaGonder(k.email)} className="btn-ghost" style={{ padding: "5px 10px", fontSize: 11.5 }}>
                          <Mail size={12} /> {sifirlamaGonderilen === k.email ? "Gönderildi ✓" : "Şifre Sıfırlama Gönder"}
                        </button>
                      )}
                    </td>
                    <td><button onClick={() => listedenKaldir(k.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        <div style={{ padding: "12px 20px", fontSize: 11.5, color: "#6b7178", borderTop: "1px solid #2a4b52" }}>
          Not: "Sil" butonu kişiyi sadece bu listeden kaldırır, giriş yapma hakkını iptal etmez. Şifreli hesap girişini tamamen kapatmak için Firebase Console → Authentication → Users sekmesinden o hesabı devre dışı bırakman gerekir; Google izinlerinde listeden silmek girişi anında keser.
        </div>
      </div>
    </div>
  );
}

// ---------- Depo Stok Listesi ----------
function DepoStokKart({ depoStok, kullanici }) {
  const [form, setForm] = useState({ stokKodu: "", stokAdi: "", birim: "Adet", anaGrupKodu: "", anaGrupAdi: "", altGrupKodu: "", altGrupAdi: "", miktar: "" });
  const [msg, setMsg] = useState("");
  const [msgTip, setMsgTip] = useState("bilgi");
  const [arama, setArama] = useState("");
  const [altGrupFiltre, setAltGrupFiltre] = useState("");
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  const [iceMsg, setIceMsg] = useState("");
  const dosyaRef = useRef(null);
  const set = (k) => (e) => setForm((s) => ({ ...s, [k]: e.target.value }));

  // ---- Toplu miktar ayarlama (tüm kalemleri belirli bir sayıya çeker) ----
  const [topluAyarDeger, setTopluAyarDeger] = useState("");
  const [topluAyarCalisiyor, setTopluAyarCalisiyor] = useState(false);
  const [topluAyarMsg, setTopluAyarMsg] = useState("");

  const topluMiktarAyarla = async () => {
    const deger = Number(topluAyarDeger);
    if (topluAyarDeger === "" || isNaN(deger)) { setTopluAyarMsg("Geçerli bir sayı gir."); setTimeout(() => setTopluAyarMsg(""), 2500); return; }
    if (!window.confirm(`TÜM stok kalemlerinin (${depoStok.length} adet) miktarı ${deger} olarak ayarlanacak. Bu işlem geri alınamaz. Emin misin?`)) return;
    setTopluAyarCalisiyor(true);
    setTopluAyarMsg("Başlıyor…");
    const PARCA = 400;
    let yapilan = 0;
    for (let i = 0; i < depoStok.length; i += PARCA) {
      const dilim = depoStok.slice(i, i + PARCA);
      const batch = writeBatch(db);
      dilim.forEach((s) => batch.update(doc(db, "depo_stok", s.id), { miktar: deger, guncellemeTarihi: Date.now() }));
      await batch.commit();
      yapilan += dilim.length;
      setTopluAyarMsg(`${yapilan} / ${depoStok.length} kalem güncellendi…`);
    }
    setTopluAyarCalisiyor(false);
    setTopluAyarDeger("");
    setTopluAyarMsg(`Tamamlandı — ${depoStok.length} kalemin miktarı ${deger} olarak ayarlandı.`);
    setTimeout(() => setTopluAyarMsg(""), 6000);
  };

  // ---- Listede birden fazla kalemin miktarını tek tek düzenleyip toplu kaydetme ----
  const [duzeltmeler, setDuzeltmeler] = useState({}); // { [stokId]: "yeni miktar" }
  const [duzeltmeKaydediliyor, setDuzeltmeKaydediliyor] = useState(false);
  const [duzeltmeMsg, setDuzeltmeMsg] = useState("");
  const duzeltmeSayisi = Object.keys(duzeltmeler).length;

  const duzeltmeDegistir = (id, deger) => setDuzeltmeler((s) => ({ ...s, [id]: deger }));

  const duzeltmeleriKaydet = async () => {
    const girdiler = Object.entries(duzeltmeler).filter(([, v]) => v !== "" && !isNaN(Number(v)));
    if (girdiler.length === 0) { setDuzeltmeMsg("Değiştirilmiş bir kalem yok."); setTimeout(() => setDuzeltmeMsg(""), 2500); return; }
    setDuzeltmeKaydediliyor(true);
    const PARCA = 400;
    let yapilan = 0;
    for (let i = 0; i < girdiler.length; i += PARCA) {
      const dilim = girdiler.slice(i, i + PARCA);
      const batch = writeBatch(db);
      dilim.forEach(([id, v]) => batch.update(doc(db, "depo_stok", id), { miktar: Number(v), guncellemeTarihi: Date.now() }));
      await batch.commit();
      yapilan += dilim.length;
      setDuzeltmeMsg(`${yapilan} / ${girdiler.length} kalem kaydediliyor…`);
    }
    setDuzeltmeler({});
    setDuzeltmeKaydediliyor(false);
    setDuzeltmeMsg(`${girdiler.length} kalemin miktarı güncellendi.`);
    setTimeout(() => setDuzeltmeMsg(""), 4000);
  };

  const anaGruplar = useMemo(() => {
    const map = new Map();
    depoStok.forEach((s) => { if (s.anaGrupAdi) map.set(s.anaGrupKodu || s.anaGrupAdi, s.anaGrupAdi); });
    return [...map.entries()];
  }, [depoStok]);
  const anaGrupAdlariListe = [...new Set(depoStok.map((s) => s.anaGrupAdi).filter(Boolean))];
  const altGrupAdlariListe = [...new Set(depoStok.map((s) => s.altGrupAdi).filter(Boolean))];

  useEffect(() => {
    if (!iceAktariliyor) return;
    const uyar = (e) => { e.preventDefault(); e.returnValue = ""; };
    window.addEventListener("beforeunload", uyar);
    return () => window.removeEventListener("beforeunload", uyar);
  }, [iceAktariliyor]);

  const ekle = async () => {
    if (!form.stokKodu || !form.stokAdi) {
      setMsgTip("hata"); setMsg("Stok Kodu ve Stok Adı zorunlu.");
      return;
    }
    if (depoStok.some((s) => s.stokKodu.toLowerCase() === form.stokKodu.trim().toLowerCase())) {
      setMsgTip("hata"); setMsg("Bu stok kodu zaten var. Miktar eklemek için \"Depo Giriş\" ekranını kullan.");
      return;
    }
    await addDoc(collection(db, "depo_stok"), {
      stokKodu: form.stokKodu.trim(), stokAdi: form.stokAdi.trim(), birim: form.birim || "Adet",
      anaGrupKodu: form.anaGrupKodu.trim(), anaGrupAdi: form.anaGrupAdi.trim(),
      altGrupKodu: form.altGrupKodu.trim(), altGrupAdi: form.altGrupAdi.trim(),
      miktar: Number(form.miktar) || 0,
      guncellemeTarihi: Date.now(),
    });
    setForm({ stokKodu: "", stokAdi: "", birim: "Adet", anaGrupKodu: "", anaGrupAdi: "", altGrupKodu: "", altGrupAdi: "", miktar: "" });
    setMsgTip("bilgi"); setMsg("Yeni stok kartı oluşturuldu.");
    setTimeout(() => setMsg(""), 3000);
  };

  const sil = async (id) => {
    if (!window.confirm("Bu stok kartı silinecek (hareket geçmişi silinmez). Emin misiniz?")) return;
    await deleteDoc(doc(db, "depo_stok", id));
  };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true);
    setIceMsg("");
    try {
      const kayitlar = await excelDenDepoOku(dosya);
      if (kayitlar.length === 0) {
        setIceMsg("Dosyada geçerli satır bulunamadı.");
      } else {
        const mevcutKodlar = new Set(depoStok.map((s) => s.stokKodu.toLowerCase()));
        const yeniler = kayitlar.filter((k) => k.stokKodu && !mevcutKodlar.has(k.stokKodu.toLowerCase()));
        const veriler = yeniler.map((k) => ({ ...k, guncellemeTarihi: Date.now() }));
        const { basarili, basarisiz } = await guvenliTopluYaz("depo_stok", veriler, (yapilan, toplam, hatali) => {
          setIceMsg(`${yapilan} / ${toplam} kayıt işlendi${hatali > 0 ? ` (${hatali} tanesi tekrar deneniyor)` : ""}…`);
        });
        // Başlangıç miktarı 0'dan büyük olan kalemler için "giriş" hareketi de oluştur,
        // böylece Stok Hareketleri / Son Girişler geçmişinde de görünürler.
        const girisliKalemler = yeniler.filter((k) => (Number(k.miktar) || 0) > 0).map((k) => ({
          stokKodu: k.stokKodu, stokAdi: k.stokAdi, tip: "giris", miktar: Number(k.miktar) || 0,
          oncekiMiktar: 0, sonrakiMiktar: Number(k.miktar) || 0, birim: k.birim || "Adet",
          aciklama: "İlk envanter (Excel içe aktarma)", kullanici: kullanici?.email || "—", tarih: Date.now(),
        }));
        if (girisliKalemler.length > 0) await guvenliTopluYaz("depo_hareketler", girisliKalemler);
        const atlanan = kayitlar.length - yeniler.length;
        let sonMesaj = `${basarili} yeni stok kartı eklendi`;
        if (atlanan > 0) sonMesaj += `, ${atlanan} tanesi zaten vardı (atlandı)`;
        if (basarisiz > 0) sonMesaj += `, ${basarisiz} tanesi eklenemedi`;
        setIceMsg(sonMesaj + ".");
      }
    } catch (err) {
      console.error(err);
      setIceMsg("İçe aktarma sırasında hata oluştu: " + (err?.message || "bilinmeyen hata"));
    }
    setIceAktariliyor(false);
    e.target.value = "";
    setTimeout(() => setIceMsg(""), 8000);
  };

  const disaAktar = () => {
    excelIndir(
      depoStok.map((s) => ({
        "STOK KODU": s.stokKodu, "STOK İSMİ": s.stokAdi, "BİRİM": s.birim || "Adet",
        "ANA GRUP KODU": s.anaGrupKodu || "", "ANA GRUP İSMİ": s.anaGrupAdi || "",
        "ALT GRUP KODU": s.altGrupKodu || "", "ALT GRUP İSMİ": s.altGrupAdi || "",
        "MİKTAR": s.miktar,
      })),
      "depo-stok-listesi.xlsx", "Depo Stok"
    );
  };

  const filtrelenmis = useMemo(() => {
    let liste = depoStok;
    if (altGrupFiltre) liste = liste.filter((s) => s.altGrupAdi === altGrupFiltre);
    if (!arama.trim()) return liste;
    const q = arama.trim().toLowerCase();
    return liste.filter((s) => s.stokKodu.toLowerCase().includes(q) || s.stokAdi.toLowerCase().includes(q) || (s.altGrupAdi || "").toLowerCase().includes(q));
  }, [depoStok, arama, altGrupFiltre]);

  const toplamKalem = depoStok.length;
  const dusukStok = depoStok.filter((s) => s.miktar <= 0).length;

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Yeni Stok Kartı Oluştur</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button
              className="btn-ghost"
              onClick={() => sablonIndir(
                ["STOK KODU", "STOK İSMİ", "BİRİM", "ANA GRUP KODU", "ANA GRUP İSMİ", "ALT GRUP KODU", "ALT GRUP İSMİ", "MİKTAR"],
                [["BTK-DDK-0001", "DELME KAFASI- Ø19,70 500000111", "Adet", "KSC", "KESİCİ TAKIM", "KSC-DEL", "KESİCİ TAKIM - DELİK DELME ELMASLARI", "0"]],
                "depo-stok-sablonu.xlsx", "Şablon"
              )}
            >
              <FileDown size={14} /> Excel Şablonu İndir
            </button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}>
              <Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excel'den İçe Aktar"}
            </button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excel'e Aktar</button>
          </div>
        </div>
        <div style={{ fontSize: 12, color: "#6b7178", marginBottom: 16 }}>Bu ekran sadece yeni stok kartı (kimlik) tanımlamak içindir. Mevcut bir stoğa miktar eklemek/düşmek için "Depo Giriş" / "Depo Çıkış" ekranlarını kullan.</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div><label className="field-label">Stok Kodu</label><input className="input" placeholder="Örn: BTK-DDK-0001" value={form.stokKodu} onChange={set("stokKodu")} /></div>
          <div style={{ gridColumn: "span 2" }}><label className="field-label">Stok İsmi</label><input className="input" placeholder="Malzeme / parça adı" value={form.stokAdi} onChange={set("stokAdi")} /></div>
          <div>
            <label className="field-label">Birim</label>
            <select className="input" value={form.birim} onChange={set("birim")}>
              <option>Adet</option><option>Kg</option><option>KİLOGRAM</option><option>Ton</option><option>MT</option><option>SANDIK</option><option>Metre</option><option>Litre</option><option>Kutu</option><option>Paket</option>
            </select>
          </div>
          <div>
            <label className="field-label">Ana Grup Kodu</label>
            <input className="input" list="ana-grup-kodlari" placeholder="Örn: KSC" value={form.anaGrupKodu} onChange={set("anaGrupKodu")} />
            <datalist id="ana-grup-kodlari">{anaGruplar.map(([kod]) => <option key={kod} value={kod} />)}</datalist>
          </div>
          <div>
            <label className="field-label">Ana Grup İsmi</label>
            <input className="input" list="ana-grup-isimleri" placeholder="Örn: KESİCİ TAKIM" value={form.anaGrupAdi} onChange={set("anaGrupAdi")} />
            <datalist id="ana-grup-isimleri">{anaGrupAdlariListe.map((v) => <option key={v} value={v} />)}</datalist>
          </div>
          <div>
            <label className="field-label">Alt Grup Kodu</label>
            <input className="input" placeholder="Örn: KSC-DEL" value={form.altGrupKodu} onChange={set("altGrupKodu")} />
          </div>
          <div>
            <label className="field-label">Alt Grup İsmi</label>
            <input className="input" list="alt-grup-isimleri" placeholder="Örn: KESİCİ TAKIM - DELİK DELME ELMASLARI" value={form.altGrupAdi} onChange={set("altGrupAdi")} />
            <datalist id="alt-grup-isimleri">{altGrupAdlariListe.map((v) => <option key={v} value={v} />)}</datalist>
          </div>
          <div><label className="field-label">Başlangıç Miktarı</label><input className="input" type="number" min="0" step="0.01" placeholder="0" value={form.miktar} onChange={set("miktar")} /></div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button onClick={ekle} style={{ display: "flex", alignItems: "center", gap: 7, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
            <Plus size={16} /> Stok Kartı Oluştur
          </button>
        </div>
        {msg && (
          <div style={{ marginTop: 14, fontSize: 12.5, color: msgTip === "hata" ? "#e07a6b" : "#2dd4bf", background: msgTip === "hata" ? "#3a1f1f" : "#113330", border: `1px solid ${msgTip === "hata" ? "#5a2a2a" : "#1f4d47"}`, borderRadius: 7, padding: "9px 12px" }}>{msg}</div>
        )}
        {iceMsg && (
          <div style={{ marginTop: 10, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{iceMsg}</div>
        )}
        <div style={{ fontSize: 12, color: "#6b7178", marginTop: 10 }}>
          Excel'den içe aktarırken başlık satırındaki isimlere bakılır (sütun sırası veya araya başka sütun eklenmiş olması önemli değil): Stok Kodu, Stok İsmi, Birim, Ana Grup Kodu, Ana Grup İsmi, Alt Grup Kodu, Alt Grup İsmi, Miktar (opsiyonel). Aynı stok kodu zaten varsa satır atlanır (mükerrer eklenmez).
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Toplam Kalem" value={toplamKalem.toLocaleString("tr-TR")} />
        <Stat label="Stokta Biten / Eksi Kalem" value={dusukStok} highlight={dusukStok > 0} />
        <Stat label="Alt Grup Sayısı" value={altGrupAdlariListe.length} />
      </div>

      <div className="card" style={{ padding: 20, borderColor: "#5a4a1f" }}>
        <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 4 }}>Toplu Miktar Ayarla (Tüm Kalemler)</div>
        <div style={{ fontSize: 12, color: "#6b7178", marginBottom: 12 }}>Listedeki <b>tüm {depoStok.length} stok kartının</b> miktarını girdiğin sayıya eşitler (mevcut değerin üzerine eklemez, doğrudan değiştirir). Geri alınamaz.</div>
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <input className="input" style={{ maxWidth: 140 }} type="number" step="0.01" placeholder="Örn: 10" value={topluAyarDeger} onChange={(e) => setTopluAyarDeger(e.target.value)} />
          <button onClick={topluMiktarAyarla} disabled={topluAyarCalisiyor} style={{ background: "#e8a33d", color: "#142a30", border: "none", borderRadius: 7, padding: "9px 16px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
            {topluAyarCalisiyor ? "Uygulanıyor…" : "Tüm Kalemlere Uygula"}
          </button>
        </div>
        {topluAyarMsg && <div style={{ marginTop: 10, fontSize: 12.5, color: "#2dd4bf" }}>{topluAyarMsg}</div>}
      </div>

      <div className="card" style={{ padding: "12px 16px" }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <div style={{ position: "relative", flex: "2 1 240px" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Stok kodu, ismi veya alt grup ara…" value={arama} onChange={(e) => setArama(e.target.value)} />
          </div>
          <select className="input" style={{ flex: "1 1 180px" }} value={altGrupFiltre} onChange={(e) => setAltGrupFiltre(e.target.value)}>
            <option value="">Tüm Alt Gruplar</option>
            {altGrupAdlariListe.map((v) => <option key={v} value={v}>{v}</option>)}
          </select>
        </div>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontWeight: 700, fontSize: 14 }}>Stok Listesi ({filtrelenmis.length.toLocaleString("tr-TR")})</div>
            <div style={{ fontSize: 11.5, color: "#6b7178", marginTop: 2 }}>Miktar hücresine tıklayıp doğrudan değiştirebilirsin, birden fazla satırı düzenleyip tek seferde kaydet.</div>
          </div>
          {duzeltmeSayisi > 0 && (
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              {duzeltmeMsg && <span style={{ fontSize: 12, color: "#2dd4bf" }}>{duzeltmeMsg}</span>}
              <button onClick={duzeltmeleriKaydet} disabled={duzeltmeKaydediliyor} style={{ background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "8px 14px", fontWeight: 700, fontSize: 12.5, cursor: "pointer" }}>
                {duzeltmeKaydediliyor ? "Kaydediliyor…" : `Değişiklikleri Kaydet (${duzeltmeSayisi})`}
              </button>
              <button onClick={() => setDuzeltmeler({})} className="btn-ghost" style={{ padding: "8px 12px", fontSize: 12.5 }}>Vazgeç</button>
            </div>
          )}
        </div>
        <div style={{ overflowX: "auto", maxHeight: 640, overflowY: "auto" }}>
          <table>
            <thead><tr><th>Stok Kodu</th><th>Stok İsmi</th><th>Ana Grup</th><th>Alt Grup</th><th>Miktar</th><th>Birim</th><th></th></tr></thead>
            <tbody>
              {filtrelenmis.length === 0 && <tr><td colSpan={7} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Henüz stok kartı yok.</td></tr>}
              {filtrelenmis.slice(0, 500).map((s) => {
                const duzenlendi = duzeltmeler[s.id] !== undefined;
                return (
                  <tr key={s.id} style={duzenlendi ? { background: "#16232a" } : undefined}>
                    <td style={{ fontFamily: "monospace" }}>{s.stokKodu}</td>
                    <td>{s.stokAdi}</td>
                    <td style={{ fontSize: 12 }}>{s.anaGrupAdi || "—"}</td>
                    <td style={{ fontSize: 12 }}>{s.altGrupAdi || "—"}</td>
                    <td>
                      <input
                        className="input" type="number" step="0.01"
                        style={{ padding: "5px 8px", fontSize: 12.5, width: 90, fontFamily: "monospace", fontWeight: 700, color: duzenlendi ? "#e8a33d" : (s.miktar <= 0 ? "#e07a6b" : "#2dd4bf"), borderColor: duzenlendi ? "#e8a33d" : undefined }}
                        value={duzeltmeler[s.id] !== undefined ? duzeltmeler[s.id] : s.miktar}
                        onChange={(e) => duzeltmeDegistir(s.id, e.target.value)}
                      />
                    </td>
                    <td>{s.birim || "Adet"}</td>
                    <td><button onClick={() => sil(s.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {filtrelenmis.length > 500 && <div style={{ padding: 12, textAlign: "center", fontSize: 12, color: "#6b7178" }}>İlk 500 kalem gösteriliyor — daha dar bir aralık görmek için arama/filtre kullan.</div>}
        </div>
      </div>
    </div>
  );
}

// ---------- Depo Stok - Depo Giriş ----------
function DepoGiris({ depoStok, kullanici, depoHareketler }) {
  const [stokId, setStokId] = useState("");
  const [miktar, setMiktar] = useState("");
  const [tarih, setTarih] = useState(todayISO());
  const [aciklama, setAciklama] = useState("");
  const [msg, setMsg] = useState("");
  const [msgTip, setMsgTip] = useState("bilgi");
  const [arama, setArama] = useState("");
  const [altGrupFiltre, setAltGrupFiltre] = useState("");
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  const [iceMsg, setIceMsg] = useState("");
  const dosyaRef = useRef(null);

  const secilenStok = depoStok.find((s) => s.id === stokId);
  const altGrupAdlariListe = [...new Set(depoStok.map((s) => s.altGrupAdi).filter(Boolean))];

  const filtrelenmisStok = useMemo(() => {
    let liste = depoStok;
    if (altGrupFiltre) liste = liste.filter((s) => s.altGrupAdi === altGrupFiltre);
    if (!arama.trim()) return liste.slice(0, 300);
    const q = arama.trim().toLowerCase();
    return liste.filter((s) => s.stokKodu.toLowerCase().includes(q) || s.stokAdi.toLowerCase().includes(q)).slice(0, 300);
  }, [depoStok, arama, altGrupFiltre]);

  // ---- Toplu stok girişi ----
  const [topluMiktarlar, setTopluMiktarlar] = useState({});
  const [topluTarih, setTopluTarih] = useState(todayISO());
  const [topluMsg, setTopluMsg] = useState("");
  const [topluKaydediliyor, setTopluKaydediliyor] = useState(false);

  const topluMiktarDegistir = (id, deger) => setTopluMiktarlar((s) => ({ ...s, [id]: deger }));

  const topluKaydet = async () => {
    const girisler = Object.entries(topluMiktarlar).filter(([, v]) => Number(v) > 0);
    if (girisler.length === 0) { setTopluMsg("En az bir kalem için miktar gir."); setTimeout(() => setTopluMsg(""), 2500); return; }
    setTopluKaydediliyor(true);
    let basarili = 0;
    for (const [id, mikStr] of girisler) {
      const stok = depoStok.find((s) => s.id === id);
      if (!stok) continue;
      const mik = Number(mikStr);
      const oncekiMiktar = stok.miktar;
      const batch = writeBatch(db);
      batch.update(doc(db, "depo_stok", id), { miktar: increment(mik), guncellemeTarihi: Date.now() });
      const hareketRef = doc(collection(db, "depo_hareketler"));
      batch.set(hareketRef, {
        stokKodu: stok.stokKodu, stokAdi: stok.stokAdi, tip: "giris", miktar: mik,
        oncekiMiktar, sonrakiMiktar: oncekiMiktar + mik, birim: stok.birim || "Adet",
        aciklama: "Toplu giriş", kullanici: kullanici?.email || "—", tarih: new Date(topluTarih).getTime() || Date.now(),
      });
      await batch.commit();
      basarili++;
    }
    setTopluMiktarlar({});
    setTopluKaydediliyor(false);
    setTopluMsg(`${basarili} kalem için giriş kaydedildi.`);
    setTimeout(() => setTopluMsg(""), 3500);
  };

  const girisYap = async () => {
    if (!stokId) { setMsgTip("hata"); setMsg("Stok kodu / adı seç."); setTimeout(() => setMsg(""), 2500); return; }
    const mik = Number(miktar);
    if (!mik || mik <= 0) { setMsgTip("hata"); setMsg("Geçerli bir miktar gir."); setTimeout(() => setMsg(""), 2500); return; }
    const oncekiMiktar = secilenStok.miktar;
    const sonrakiMiktar = oncekiMiktar + mik;
    const batch = writeBatch(db);
    batch.update(doc(db, "depo_stok", stokId), { miktar: increment(mik), guncellemeTarihi: Date.now() });
    const hareketRef = doc(collection(db, "depo_hareketler"));
    batch.set(hareketRef, {
      stokKodu: secilenStok.stokKodu, stokAdi: secilenStok.stokAdi, tip: "giris", miktar: mik,
      oncekiMiktar, sonrakiMiktar, birim: secilenStok.birim || "Adet",
      aciklama: aciklama.trim(), kullanici: kullanici?.email || "—", tarih: new Date(tarih).getTime() || Date.now(),
    });
    await batch.commit();
    setMiktar(""); setAciklama("");
    setMsgTip("bilgi"); setMsg(`${secilenStok.stokAdi} için ${mik} ${secilenStok.birim || "Adet"} giriş yapıldı.`);
    setTimeout(() => setMsg(""), 3500);
  };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true); setIceMsg("");
    try {
      const rows = await dosyaOku(dosya);
      let baslangic = 0;
      const ilkSatir = (rows[0] || []).map((v) => String(v || "").trim().toLowerCase());
      if (ilkSatir[0] && ilkSatir[0].includes("stok")) baslangic = 1;
      let basarili = 0, atlanan = 0;
      for (let i = baslangic; i < rows.length; i++) {
        const r = rows[i] || [];
        const stokKodu = String(r[0] || "").trim();
        const mik = sayiAyristir(r[1]);
        const tarihStr = String(r[2] || "").trim();
        const aciklamaStr = String(r[3] || "").trim();
        const stok = depoStok.find((s) => s.stokKodu.toLowerCase() === stokKodu.toLowerCase());
        if (!stok || !mik) { atlanan++; continue; }
        const oncekiMiktar = stok.miktar;
        const batch = writeBatch(db);
        batch.update(doc(db, "depo_stok", stok.id), { miktar: increment(mik), guncellemeTarihi: Date.now() });
        const hareketRef = doc(collection(db, "depo_hareketler"));
        batch.set(hareketRef, {
          stokKodu: stok.stokKodu, stokAdi: stok.stokAdi, tip: "giris", miktar: mik,
          oncekiMiktar, sonrakiMiktar: oncekiMiktar + mik, birim: stok.birim || "Adet",
          aciklama: aciklamaStr, kullanici: kullanici?.email || "—",
          tarih: tarihStr ? new Date(tarihStr).getTime() : Date.now(),
        });
        await batch.commit();
        basarili++;
      }
      setIceMsg(`${basarili} giriş işlendi${atlanan > 0 ? `, ${atlanan} satır atlandı (stok kodu bulunamadı ya da eksik bilgi)` : ""}.`);
    } catch (err) {
      console.error(err);
      setIceMsg("Hata: " + (err?.message || "bilinmeyen hata"));
    }
    setIceAktariliyor(false); e.target.value = ""; setTimeout(() => setIceMsg(""), 7000);
  };

  const sonGirisler = useMemo(() => {
    return [...depoHareketler].filter((h) => h.tip === "giris").sort((a, b) => (b.tarih || 0) - (a.tarih || 0)).slice(0, 15);
  }, [depoHareketler]);

  const disaAktar = () => excelIndir(
    sonGirisler.map((h) => ({
      Tarih: h.tarih ? new Date(h.tarih).toLocaleString("tr-TR") : "", "Stok Kodu": h.stokKodu, "Stok Adı": h.stokAdi,
      Miktar: h.miktar, Birim: h.birim || "Adet", Kullanıcı: h.kullanici, Açıklama: h.aciklama,
    })), "stok-girisleri.xlsx", "Girişler"
  );

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Depoya Stok Girişi</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn-ghost" onClick={() => sablonIndir(["Stok Kodu", "Miktar", "Tarih", "Açıklama"], [["ORN-001", 50, "2026-01-15", ""]], "depo-giris-sablonu.xlsx", "Şablon")}><FileDown size={14} /> Şablon İndir</button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}><Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excel'den İçe Aktar"}</button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excele Aktar</button>
          </div>
        </div>
        {iceMsg && <div style={{ marginBottom: 14, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{iceMsg}</div>}
        <div style={{ fontSize: 12, color: "#6b7178", marginBottom: 16 }}>Var olan bir stok kartına gelen malzemeyi/ürünü ekle.</div>

        <div style={{ marginBottom: 14 }}>
          <label className="field-label">Alt Grup (opsiyonel filtre)</label>
          <select className="input" value={altGrupFiltre} onChange={(e) => setAltGrupFiltre(e.target.value)}>
            <option value="">Tüm Alt Gruplar</option>
            {altGrupAdlariListe.map((v) => <option key={v} value={v}>{v}</option>)}
          </select>
        </div>
        <div style={{ marginBottom: 14, position: "relative" }}>
          <label className="field-label">Stok Kodu / Adı Ara</label>
          <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: 34 }} />
          <input className="input" style={{ paddingLeft: 30 }} placeholder="Ara…" value={arama} onChange={(e) => { setArama(e.target.value); setStokId(""); }} />
        </div>

        {(arama.trim() || altGrupFiltre) && (
          <div style={{ marginBottom: 14, border: "1px solid #2a4b52", borderRadius: 8, maxHeight: 260, overflowY: "auto" }}>
            {filtrelenmisStok.length === 0 && <div style={{ padding: 14, fontSize: 12.5, color: "#6b7178", textAlign: "center" }}>Eşleşen stok bulunamadı.</div>}
            {filtrelenmisStok.slice(0, 50).map((s) => (
              <button
                key={s.id}
                onClick={() => setStokId(s.id)}
                style={{
                  display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%",
                  padding: "10px 14px", background: stokId === s.id ? "#113330" : "transparent",
                  border: "none", borderBottom: "1px solid #223b42", cursor: "pointer", textAlign: "left",
                }}
              >
                <div>
                  <div style={{ fontSize: 13, color: stokId === s.id ? "#2dd4bf" : "#e7e5e0", fontWeight: stokId === s.id ? 700 : 500 }}>{s.stokAdi}</div>
                  <div style={{ fontSize: 11.5, color: "#6b7178", fontFamily: "monospace" }}>{s.stokKodu}{s.altGrupAdi ? ` · ${s.altGrupAdi}` : ""}</div>
                </div>
                <div style={{ fontFamily: "monospace", fontSize: 12.5, color: s.miktar <= 0 ? "#e07a6b" : "#8b929a", flexShrink: 0, marginLeft: 10 }}>{s.miktar} {s.birim || "Adet"}</div>
              </button>
            ))}
            {filtrelenmisStok.length > 50 && <div style={{ padding: 10, fontSize: 11.5, color: "#6b7178", textAlign: "center" }}>İlk 50 sonuç gösteriliyor, daraltmak için ara.</div>}
          </div>
        )}

        {secilenStok && (
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "#113330", border: "1px solid #1f4d47", borderRadius: 8, padding: "10px 14px", marginBottom: 14 }}>
            <div style={{ fontSize: 13 }}>
              Seçili: <b style={{ color: "#2dd4bf" }}>{secilenStok.stokAdi}</b> <span style={{ color: "#6b7178", fontFamily: "monospace" }}>({secilenStok.stokKodu})</span> — mevcut: {secilenStok.miktar} {secilenStok.birim || "Adet"}
            </div>
            <button onClick={() => setStokId("")} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", fontSize: 12 }}>Değiştir ✕</button>
          </div>
        )}

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 14 }}>
          <div style={{ gridColumn: "span 2" }}>
            <label className="field-label">Stok Adı</label>
            <input className="input" value={secilenStok ? secilenStok.stokAdi : ""} placeholder="Yukarıdan bir stok seç" disabled style={{ opacity: secilenStok ? 1 : 0.6 }} />
          </div>
          <div><label className="field-label">Miktar</label><input className="input" type="number" min="0" step="0.01" value={miktar} onChange={(e) => setMiktar(e.target.value)} /></div>
          <div><label className="field-label">Tarih</label><input className="input" type="date" value={tarih} onChange={(e) => setTarih(e.target.value)} /></div>
          <div style={{ gridColumn: "span 2" }}><label className="field-label">Açıklama (opsiyonel)</label><input className="input" value={aciklama} onChange={(e) => setAciklama(e.target.value)} /></div>
        </div>

        <button onClick={girisYap} style={{ background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 18px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
          + Girişi Kaydet
        </button>
        {msg && (
          <div style={{ marginTop: 14, fontSize: 12.5, color: msgTip === "hata" ? "#e07a6b" : "#2dd4bf", background: msgTip === "hata" ? "#3a1f1f" : "#113330", border: `1px solid ${msgTip === "hata" ? "#5a2a2a" : "#1f4d47"}`, borderRadius: 7, padding: "9px 12px" }}>{msg}</div>
        )}
        {depoStok.length === 0 && <div style={{ marginTop: 10, fontSize: 12, color: "#e8a33d" }}>Önce "Stok Kartı Oluştur" ekranından bir stok kartı ekle.</div>}
      </div>

      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>Toplu Stok Girişi</div>
        <div style={{ fontSize: 12, color: "#6b7178", marginBottom: 16 }}>Yukarıdaki Alt Grup / arama ile listeyi daralt, birden fazla kalemin miktarını aynı anda yazıp tek seferde kaydet.</div>
        {(arama.trim() || altGrupFiltre) ? (
          <>
            <div style={{ border: "1px solid #2a4b52", borderRadius: 8, maxHeight: 340, overflowY: "auto", marginBottom: 14 }}>
              <table>
                <thead><tr><th>Stok Adı</th><th>Stok Kodu</th><th>Mevcut</th><th style={{ width: 120 }}>Giriş Miktarı</th></tr></thead>
                <tbody>
                  {filtrelenmisStok.length === 0 && <tr><td colSpan={4} style={{ color: "#6b7178", textAlign: "center", padding: 18 }}>Eşleşen stok bulunamadı.</td></tr>}
                  {filtrelenmisStok.slice(0, 100).map((s) => (
                    <tr key={s.id}>
                      <td style={{ fontSize: 13 }}>{s.stokAdi}</td>
                      <td style={{ fontFamily: "monospace", fontSize: 12 }}>{s.stokKodu}</td>
                      <td style={{ fontFamily: "monospace", fontSize: 12, color: s.miktar <= 0 ? "#e07a6b" : "#8b929a" }}>{s.miktar} {s.birim || "Adet"}</td>
                      <td>
                        <input
                          className="input" type="number" min="0" step="0.01" placeholder="0"
                          style={{ padding: "6px 8px", fontSize: 12.5 }}
                          value={topluMiktarlar[s.id] || ""}
                          onChange={(e) => topluMiktarDegistir(s.id, e.target.value)}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {filtrelenmisStok.length > 100 && <div style={{ padding: 10, fontSize: 11.5, color: "#6b7178", textAlign: "center" }}>İlk 100 sonuç gösteriliyor, daraltmak için ara.</div>}
            </div>
            <div style={{ display: "flex", alignItems: "flex-end", gap: 14, flexWrap: "wrap" }}>
              <div style={{ maxWidth: 200 }}>
                <label className="field-label">Tarih (tüm kalemler için)</label>
                <input className="input" type="date" value={topluTarih} onChange={(e) => setTopluTarih(e.target.value)} />
              </div>
              <button onClick={topluKaydet} disabled={topluKaydediliyor} style={{ background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "10px 18px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
                {topluKaydediliyor ? "Kaydediliyor…" : "+ Toplu Girişi Kaydet"}
              </button>
            </div>
            {topluMsg && <div style={{ marginTop: 14, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{topluMsg}</div>}
          </>
        ) : (
          <div style={{ fontSize: 12.5, color: "#6b7178" }}>Listeyi görmek için yukarıdan bir Alt Grup seç ya da arama yap.</div>
        )}
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Son Girişler</div>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr><th>Tarih</th><th>Stok</th><th>Miktar</th><th>Kullanıcı</th><th>Açıklama</th></tr></thead>
            <tbody>
              {sonGirisler.length === 0 && <tr><td colSpan={5} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Henüz kayıtlı giriş yok.</td></tr>}
              {sonGirisler.map((h) => (
                <tr key={h.id}>
                  <td style={{ fontFamily: "monospace", whiteSpace: "nowrap" }}>{h.tarih ? new Date(h.tarih).toLocaleString("tr-TR") : "—"}</td>
                  <td>{h.stokAdi} <span style={{ color: "#6b7178", fontSize: 11.5 }}>({h.stokKodu})</span></td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#2dd4bf" }}>{h.miktar} {h.birim || ""}</td>
                  <td style={{ fontSize: 12 }}>{h.kullanici}</td>
                  <td style={{ fontSize: 12.5 }}>{h.aciklama || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Depo Stok - Stok Çıkış (Hedef Makine Seçimli) ----------
function DepoStokCikis({ depoStok, machines, kullanici, depoHareketler }) {
  const [stokId, setStokId] = useState("");
  const [hedefMakine, setHedefMakine] = useState("");
  const [miktar, setMiktar] = useState("");
  const [tarih, setTarih] = useState(todayISO());
  const [aciklama, setAciklama] = useState("");
  const [msg, setMsg] = useState("");
  const [msgTip, setMsgTip] = useState("bilgi");
  const [arama, setArama] = useState("");
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  const [iceMsg, setIceMsg] = useState("");
  const dosyaRef = useRef(null);

  const secilenStok = depoStok.find((s) => s.id === stokId);
  const [altGrupFiltre, setAltGrupFiltre] = useState("");
  const altGrupAdlariListe = [...new Set(depoStok.map((s) => s.altGrupAdi).filter(Boolean))];

  const filtrelenmisStok = useMemo(() => {
    let liste = depoStok;
    if (altGrupFiltre) liste = liste.filter((s) => s.altGrupAdi === altGrupFiltre);
    if (!arama.trim()) return liste.slice(0, 300);
    const q = arama.trim().toLowerCase();
    return liste.filter((s) => s.stokKodu.toLowerCase().includes(q) || s.stokAdi.toLowerCase().includes(q)).slice(0, 300);
  }, [depoStok, arama, altGrupFiltre]);

  const cikisYap = async () => {
    if (!stokId) { setMsgTip("hata"); setMsg("Stok kodu / adı seç."); setTimeout(() => setMsg(""), 2500); return; }
    if (!hedefMakine) { setMsgTip("hata"); setMsg("Hangi makineye çıkış yapıldığını seç."); setTimeout(() => setMsg(""), 2500); return; }
    const mik = Number(miktar);
    if (!mik || mik <= 0) { setMsgTip("hata"); setMsg("Geçerli bir miktar gir."); setTimeout(() => setMsg(""), 2500); return; }
    if (mik > secilenStok.miktar) {
      if (!window.confirm(`Mevcut stok ${secilenStok.miktar} ${secilenStok.birim}. Bu çıkış stoğu eksiye düşürecek. Devam edilsin mi?`)) return;
    }
    const oncekiMiktar = secilenStok.miktar;
    const sonrakiMiktar = oncekiMiktar - mik;
    const batch = writeBatch(db);
    batch.update(doc(db, "depo_stok", stokId), { miktar: increment(-mik), guncellemeTarihi: Date.now() });
    const hareketRef = doc(collection(db, "depo_hareketler"));
    batch.set(hareketRef, {
      stokKodu: secilenStok.stokKodu, stokAdi: secilenStok.stokAdi, tip: "cikis", miktar: mik,
      oncekiMiktar, sonrakiMiktar, birim: secilenStok.birim || "Adet",
      hedefMakine, aciklama: aciklama.trim(), kullanici: kullanici?.email || "—", tarih: new Date(tarih).getTime() || Date.now(),
    });
    await batch.commit();
    setMiktar(""); setAciklama("");
    setMsgTip("bilgi"); setMsg(`${secilenStok.stokAdi} → ${hedefMakine} : ${mik} ${secilenStok.birim || "Adet"} çıkış yapıldı.`);
    setTimeout(() => setMsg(""), 3500);
  };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true); setIceMsg("");
    try {
      const rows = await dosyaOku(dosya);
      let baslangic = 0;
      const ilkSatir = (rows[0] || []).map((v) => String(v || "").trim().toLowerCase());
      if (ilkSatir[0] && ilkSatir[0].includes("stok")) baslangic = 1;
      let basarili = 0, atlanan = 0;
      for (let i = baslangic; i < rows.length; i++) {
        const r = rows[i] || [];
        const stokKodu = String(r[0] || "").trim();
        const makine = String(r[1] || "").trim();
        const mik = sayiAyristir(r[2]);
        const tarihStr = String(r[3] || "").trim();
        const aciklamaStr = String(r[4] || "").trim();
        const stok = depoStok.find((s) => s.stokKodu.toLowerCase() === stokKodu.toLowerCase());
        if (!stok || !makine || !mik) { atlanan++; continue; }
        const oncekiMiktar = stok.miktar;
        const batch = writeBatch(db);
        batch.update(doc(db, "depo_stok", stok.id), { miktar: increment(-mik), guncellemeTarihi: Date.now() });
        const hareketRef = doc(collection(db, "depo_hareketler"));
        batch.set(hareketRef, {
          stokKodu: stok.stokKodu, stokAdi: stok.stokAdi, tip: "cikis", miktar: mik,
          oncekiMiktar, sonrakiMiktar: oncekiMiktar - mik, birim: stok.birim || "Adet",
          hedefMakine: makine, aciklama: aciklamaStr, kullanici: kullanici?.email || "—",
          tarih: tarihStr ? new Date(tarihStr).getTime() : Date.now(),
        });
        await batch.commit();
        basarili++;
      }
      setIceMsg(`${basarili} çıkış işlendi${atlanan > 0 ? `, ${atlanan} satır atlandı (stok kodu bulunamadı ya da eksik bilgi)` : ""}.`);
    } catch (err) {
      console.error(err);
      setIceMsg("Hata: " + (err?.message || "bilinmeyen hata"));
    }
    setIceAktariliyor(false); e.target.value = ""; setTimeout(() => setIceMsg(""), 7000);
  };

  const sonCikislar = useMemo(() => {
    return [...depoHareketler].filter((h) => h.tip === "cikis" && h.hedefMakine).sort((a, b) => (b.tarih || 0) - (a.tarih || 0)).slice(0, 15);
  }, [depoHareketler]);

  const disaAktar = () => excelIndir(
    sonCikislar.map((h) => ({
      Tarih: h.tarih ? new Date(h.tarih).toLocaleString("tr-TR") : "", "Stok Kodu": h.stokKodu, "Stok Adı": h.stokAdi,
      "Hedef Makine": h.hedefMakine, Miktar: h.miktar, Birim: h.birim || "Adet", Kullanıcı: h.kullanici,
    })), "stok-cikislari.xlsx", "Çıkışlar"
  );

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Depodan Stok Çıkışı</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn-ghost" onClick={() => sablonIndir(["Stok Kodu", "Hedef Makine", "Miktar", "Tarih", "Açıklama"], [["STK-001", "Makine A", 5, "2026-01-15", ""]], "stok-cikis-sablonu.xlsx", "Şablon")}><FileDown size={14} /> Şablon İndir</button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}><Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excel'den İçe Aktar"}</button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excele Aktar</button>
          </div>
        </div>
        {iceMsg && <div style={{ marginBottom: 14, fontSize: 12.5, color: "#2dd4bf", background: "#113330", border: "1px solid #1f4d47", borderRadius: 7, padding: "9px 12px" }}>{iceMsg}</div>}
        <div style={{ fontSize: 12, color: "#6b7178", marginBottom: 16 }}>Verilen ürünü ve hangi makineye gittiğini seç, adedini yaz, çıkışı kaydet — ilgili stoktan otomatik düşer.</div>

        <div style={{ marginBottom: 14 }}>
          <label className="field-label">Alt Grup (opsiyonel filtre)</label>
          <select className="input" value={altGrupFiltre} onChange={(e) => setAltGrupFiltre(e.target.value)}>
            <option value="">Tüm Alt Gruplar</option>
            {altGrupAdlariListe.map((v) => <option key={v} value={v}>{v}</option>)}
          </select>
        </div>
        <div style={{ marginBottom: 14, position: "relative" }}>
          <label className="field-label">Stok Kodu / Adı Ara</label>
          <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: 34, }} />
          <input className="input" style={{ paddingLeft: 30 }} placeholder="Ara…" value={arama} onChange={(e) => { setArama(e.target.value); setStokId(""); }} />
        </div>

        {(arama.trim() || altGrupFiltre) && (
          <div style={{ marginBottom: 14, border: "1px solid #2a4b52", borderRadius: 8, maxHeight: 260, overflowY: "auto" }}>
            {filtrelenmisStok.length === 0 && <div style={{ padding: 14, fontSize: 12.5, color: "#6b7178", textAlign: "center" }}>Eşleşen stok bulunamadı.</div>}
            {filtrelenmisStok.slice(0, 50).map((s) => (
              <button
                key={s.id}
                onClick={() => setStokId(s.id)}
                style={{
                  display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%",
                  padding: "10px 14px", background: stokId === s.id ? "#113330" : "transparent",
                  border: "none", borderBottom: "1px solid #223b42", cursor: "pointer", textAlign: "left",
                }}
              >
                <div>
                  <div style={{ fontSize: 13, color: stokId === s.id ? "#2dd4bf" : "#e7e5e0", fontWeight: stokId === s.id ? 700 : 500 }}>{s.stokAdi}</div>
                  <div style={{ fontSize: 11.5, color: "#6b7178", fontFamily: "monospace" }}>{s.stokKodu}{s.altGrupAdi ? ` · ${s.altGrupAdi}` : ""}</div>
                </div>
                <div style={{ fontFamily: "monospace", fontSize: 12.5, color: s.miktar <= 0 ? "#e07a6b" : "#8b929a", flexShrink: 0, marginLeft: 10 }}>{s.miktar} {s.birim || "Adet"}</div>
              </button>
            ))}
            {filtrelenmisStok.length > 50 && <div style={{ padding: 10, fontSize: 11.5, color: "#6b7178", textAlign: "center" }}>İlk 50 sonuç gösteriliyor, daraltmak için ara.</div>}
          </div>
        )}

        {secilenStok && (
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "#113330", border: "1px solid #1f4d47", borderRadius: 8, padding: "10px 14px", marginBottom: 14 }}>
            <div style={{ fontSize: 13 }}>
              Seçili: <b style={{ color: "#2dd4bf" }}>{secilenStok.stokAdi}</b> <span style={{ color: "#6b7178", fontFamily: "monospace" }}>({secilenStok.stokKodu})</span>
            </div>
            <button onClick={() => setStokId("")} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", fontSize: 12 }}>Değiştir ✕</button>
          </div>
        )}

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 14 }}>
          <div style={{ gridColumn: "span 2" }}>
            <label className="field-label">Stok Adı</label>
            <input className="input" value={secilenStok ? secilenStok.stokAdi : ""} placeholder="Yukarıdan bir stok seç" disabled style={{ opacity: secilenStok ? 1 : 0.6 }} />
          </div>
          <div>
            <label className="field-label">Hedef Makine</label>
            <select className="input" value={hedefMakine} onChange={(e) => setHedefMakine(e.target.value)}>
              <option value="">Seçiniz</option>
              {machines.map((m) => <option key={m.id} value={m.name}>{m.name}</option>)}
            </select>
          </div>
          <div><label className="field-label">Miktar</label><input className="input" type="number" min="0" step="0.01" value={miktar} onChange={(e) => setMiktar(e.target.value)} /></div>
          <div><label className="field-label">Tarih</label><input className="input" type="date" value={tarih} onChange={(e) => setTarih(e.target.value)} /></div>
          <div style={{ gridColumn: "span 2" }}><label className="field-label">Açıklama (opsiyonel)</label><input className="input" value={aciklama} onChange={(e) => setAciklama(e.target.value)} /></div>
        </div>

        <button onClick={cikisYap} style={{ background: "#c0392b", color: "#fff", border: "none", borderRadius: 7, padding: "10px 18px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
          − Çıkışı Kaydet
        </button>
        {msg && (
          <div style={{ marginTop: 14, fontSize: 12.5, color: msgTip === "hata" ? "#e07a6b" : "#2dd4bf", background: msgTip === "hata" ? "#3a1f1f" : "#113330", border: `1px solid ${msgTip === "hata" ? "#5a2a2a" : "#1f4d47"}`, borderRadius: 7, padding: "9px 12px" }}>{msg}</div>
        )}
        {machines.length === 0 && <div style={{ marginTop: 10, fontSize: 12, color: "#e8a33d" }}>Önce "Makineler" ekranından makine ekle.</div>}
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Son Çıkışlar</div>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr><th>Tarih</th><th>Stok</th><th>Hedef Makine</th><th>Miktar</th><th>Kullanıcı</th></tr></thead>
            <tbody>
              {sonCikislar.length === 0 && <tr><td colSpan={5} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Henüz kayıtlı çıkış yok.</td></tr>}
              {sonCikislar.map((h) => (
                <tr key={h.id}>
                  <td style={{ fontFamily: "monospace", whiteSpace: "nowrap" }}>{h.tarih ? new Date(h.tarih).toLocaleString("tr-TR") : "—"}</td>
                  <td>{h.stokAdi} <span style={{ color: "#6b7178", fontSize: 11.5 }}>({h.stokKodu})</span></td>
                  <td><span className="pill">{h.hedefMakine}</span></td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#e07a6b" }}>{h.miktar} {h.birim || ""}</td>
                  <td style={{ fontSize: 12 }}>{h.kullanici}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Stok Hareketleri (Depo hareket geçmişi) ----------
function DepoHareketleri({ depoHareketler }) {
  const [f, setF] = useState({ arama: "", tip: "", stokKodu: "", baslangic: "", bitis: "" });
  const setF2 = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const donemSec = (tur) => {
    const bugun = new Date();
    if (tur === "gun") {
      const g = todayISO();
      setF((s) => ({ ...s, baslangic: g, bitis: g }));
    } else if (tur === "ay") {
      const ilkGun = new Date(bugun.getFullYear(), bugun.getMonth(), 1).toISOString().slice(0, 10);
      setF((s) => ({ ...s, baslangic: ilkGun, bitis: todayISO() }));
    } else if (tur === "yil") {
      const ilkGun = new Date(bugun.getFullYear(), 0, 1).toISOString().slice(0, 10);
      setF((s) => ({ ...s, baslangic: ilkGun, bitis: todayISO() }));
    } else {
      setF((s) => ({ ...s, baslangic: "", bitis: "" }));
    }
  };

  const stokKodlari = [...new Set(depoHareketler.map((h) => h.stokKodu).filter(Boolean))];

  const filtrelenmis = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return depoHareketler.filter((h) => {
      if (f.tip && h.tip !== f.tip) return false;
      if (f.stokKodu && h.stokKodu !== f.stokKodu) return false;
      if (f.baslangic && h.tarih && new Date(h.tarih).toISOString().slice(0, 10) < f.baslangic) return false;
      if (f.bitis && h.tarih && new Date(h.tarih).toISOString().slice(0, 10) > f.bitis) return false;
      if (q && !(
        (h.stokKodu || "").toLowerCase().includes(q) ||
        (h.stokAdi || "").toLowerCase().includes(q) ||
        (h.aciklama || "").toLowerCase().includes(q) ||
        (h.kullanici || "").toLowerCase().includes(q)
      )) return false;
      return true;
    }).sort((a, b) => (b.tarih || 0) - (a.tarih || 0));
  }, [depoHareketler, f]);

  const toplamGiris = filtrelenmis.filter((h) => h.tip === "giris").reduce((s, h) => s + (h.miktar || 0), 0);
  const toplamCikis = filtrelenmis.filter((h) => h.tip === "cikis").reduce((s, h) => s + (h.miktar || 0), 0);

  const disaAktar = () => {
    excelIndir(
      filtrelenmis.map((h) => ({
        Tarih: h.tarih ? new Date(h.tarih).toLocaleString("tr-TR") : "",
        "Stok Kodu": h.stokKodu, "Stok Adı": h.stokAdi,
        Tip: h.tip === "giris" ? "Giriş" : "Çıkış", Miktar: h.miktar, Birim: h.birim || "Adet",
        "Hedef Makine": h.hedefMakine || "", "Önceki Miktar": h.oncekiMiktar, "Sonraki Miktar": h.sonrakiMiktar,
        Kullanıcı: h.kullanici, Açıklama: h.aciklama,
      })),
      "stok-hareketleri.xlsx", "Hareketler"
    );
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Filtrele</div>
          <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excel'e Aktar</button>
        </div>
        <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
          <button className="btn-ghost" onClick={() => donemSec("gun")}>Bugün</button>
          <button className="btn-ghost" onClick={() => donemSec("ay")}>Bu Ay</button>
          <button className="btn-ghost" onClick={() => donemSec("yil")}>Bu Yıl</button>
          <button className="btn-ghost" onClick={() => donemSec("tumu")}>Tüm Zamanlar</button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1", position: "relative" }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input className="input" style={{ paddingLeft: 30 }} placeholder="Stok kodu, adı, kullanıcı, açıklama ara…" value={f.arama} onChange={setF2("arama")} />
          </div>
          <div>
            <label className="field-label">Tip</label>
            <select className="input" value={f.tip} onChange={setF2("tip")}>
              <option value="">Tümü</option>
              <option value="giris">Giriş</option>
              <option value="cikis">Çıkış</option>
            </select>
          </div>
          <div>
            <label className="field-label">Stok Kodu</label>
            <select className="input" value={f.stokKodu} onChange={setF2("stokKodu")}>
              <option value="">Tümü</option>
              {stokKodlari.map((k) => <option key={k} value={k}>{k}</option>)}
            </select>
          </div>
          <div><label className="field-label">Başlangıç</label><input className="input" type="date" value={f.baslangic} onChange={setF2("baslangic")} /></div>
          <div><label className="field-label">Bitiş</label><input className="input" type="date" value={f.bitis} onChange={setF2("bitis")} /></div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Toplam Hareket" value={filtrelenmis.length} />
        <Stat label="Toplam Giriş" value={toplamGiris.toLocaleString("tr-TR")} highlight />
        <Stat label="Toplam Çıkış" value={toplamCikis.toLocaleString("tr-TR")} />
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a4b52", fontWeight: 700, fontSize: 14 }}>Hareketler ({filtrelenmis.length})</div>
        <div style={{ overflowX: "auto", maxHeight: 600, overflowY: "auto" }}>
          <table>
            <thead><tr><th>Tarih</th><th>Stok Kodu</th><th>Stok Adı</th><th>Tip</th><th>Miktar</th><th>Hedef Makine</th><th>Önceki → Sonraki</th><th>Kullanıcı</th><th>Açıklama</th></tr></thead>
            <tbody>
              {filtrelenmis.length === 0 && <tr><td colSpan={9} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Kayıt bulunamadı.</td></tr>}
              {filtrelenmis.map((h) => (
                <tr key={h.id}>
                  <td style={{ fontFamily: "monospace", whiteSpace: "nowrap" }}>{h.tarih ? new Date(h.tarih).toLocaleString("tr-TR") : "—"}</td>
                  <td style={{ fontFamily: "monospace" }}>{h.stokKodu}</td>
                  <td>{h.stokAdi}</td>
                  <td>
                    {h.tip === "giris"
                      ? <span className="pill">+ Giriş</span>
                      : <span className="pill" style={{ background: "#3a1f1f", color: "#e07a6b", borderColor: "#5a2a2a" }}>− Çıkış</span>}
                  </td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700 }}>{h.miktar} {h.birim || ""}</td>
                  <td>{h.hedefMakine ? <span className="pill">{h.hedefMakine}</span> : "—"}</td>
                  <td style={{ fontFamily: "monospace", color: "#8b929a" }}>{h.oncekiMiktar} → {h.sonrakiMiktar}</td>
                  <td style={{ fontSize: 12 }}>{h.kullanici || "—"}</td>
                  <td style={{ fontSize: 12.5 }}>{h.aciklama || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}


function ListeYonetimi({ title, baslikCogul, koleksiyon, placeholder, items, icon: Icon }) {
  const [val, setVal] = useState("");
  const [arama, setArama] = useState("");
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  useEffect(() => {
    if (!iceAktariliyor) return;
    const uyar = (e) => { e.preventDefault(); e.returnValue = ""; };
    window.addEventListener("beforeunload", uyar);
    return () => window.removeEventListener("beforeunload", uyar);
  }, [iceAktariliyor]);
  const [mesaj, setMesaj] = useState("");
  const dosyaRef = useRef(null);

  const add = async () => {
    const name = val.trim();
    if (!name) return;
    if (items.some((i) => i.name.toLowerCase() === name.toLowerCase())) { setVal(""); return; }
    await addDoc(collection(db, koleksiyon), { name });
    setVal("");
  };
  const sil = async (id) => { await deleteDoc(doc(db, koleksiyon, id)); };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true);
    setMesaj("");
    try {
      const isimler = await excelDenIsimOku(dosya);
      const mevcut = new Set(items.map((i) => i.name.toLowerCase()));
      const yeniler = isimler.filter((n) => !mevcut.has(n.toLowerCase()));
      const veriler = yeniler.map((name) => ({ name }));
      const { basarili, basarisiz } = await guvenliTopluYaz(koleksiyon, veriler);
      let sonMesaj = `${basarili} yeni ${title.toLowerCase()} eklendi`;
      if (isimler.length - yeniler.length > 0) sonMesaj += `, ${isimler.length - yeniler.length} zaten vardı`;
      if (basarisiz > 0) sonMesaj += `, ${basarisiz} tanesi eklenemedi (bağlantı sorunu, tekrar dene)`;
      setMesaj(sonMesaj + ".");
    } catch (err) {
      console.error(err);
      setMesaj("Dosya okunamadı. .xlsx veya .csv dosyası olduğundan emin olun.");
    }
    setIceAktariliyor(false);
    e.target.value = "";
    setTimeout(() => setMesaj(""), 6000);
  };

  const disaAktar = () => {
    excelIndir(items.map((i) => ({ [title]: i.name })), `${baslikCogul.toLowerCase()}.xlsx`, baslikCogul);
  };

  const filtrelenmis = useMemo(() => {
    if (!arama.trim()) return items;
    const q = arama.trim().toLowerCase();
    return items.filter((i) => i.name.toLowerCase().includes(q));
  }, [items, arama]);

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15, display: "flex", alignItems: "center", gap: 8 }}>
            <Icon size={17} color="#2dd4bf" /> {baslikCogul} Listesi
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button
              className="btn-ghost"
              onClick={() => sablonIndir(
                [title],
                [[`Örnek ${title} 1`], [`Örnek ${title} 2`]],
                `${baslikCogul.toLowerCase()}-sablonu.xlsx`, baslikCogul
              )}
            >
              <FileDown size={14} /> Excel Şablonu İndir
            </button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}>
              <Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excel'den İçe Aktar"}
            </button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excel'e Aktar</button>
          </div>
        </div>

        <div style={{ display: "flex", gap: 10 }}>
          <input
            className="input"
            placeholder={placeholder}
            value={val}
            onChange={(e) => setVal(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && add()}
          />
          <button onClick={add} style={{ display: "flex", alignItems: "center", gap: 6, background: "#2dd4bf", color: "#142a30", border: "none", borderRadius: 7, padding: "0 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer", whiteSpace: "nowrap" }}>
            <Plus size={15} /> Ekle
          </button>
        </div>
        <div style={{ fontSize: 12, color: mesaj ? "#2dd4bf" : "#6b7178", marginTop: 10 }}>
          {mesaj || `Toplam ${items.length} kayıt. Excel dosyasında isimler tek sütunda alt alta olmalı (başlık satırı olabilir).`}
        </div>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "12px 16px", borderBottom: "1px solid #2a4b52", position: "relative" }}>
          <Search size={14} color="#6b7178" style={{ position: "absolute", left: 26, top: "50%", transform: "translateY(-50%)" }} />
          <input
            className="input"
            style={{ paddingLeft: 30 }}
            placeholder={`${baslikCogul} içinde ara…`}
            value={arama}
            onChange={(e) => setArama(e.target.value)}
          />
        </div>
        {filtrelenmis.length === 0 ? (
          <div style={{ color: "#6b7178", textAlign: "center", padding: 32, fontSize: 13.5 }}>
            {arama ? "Sonuç bulunamadı." : `Henüz ${baslikCogul.toLowerCase()} eklenmedi.`}
          </div>
        ) : (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))" }}>
            {filtrelenmis.map((i, idx) => (
              <div key={i.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 16px", borderBottom: "1px solid #24424a", borderRight: "1px solid #24424a" }}>
                <span style={{ fontSize: 13.5, display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontFamily: "monospace", color: "#8b929a", fontSize: 11 }}>{String(idx + 1).padStart(2, "0")}</span>
                  {i.name}
                </span>
                <button onClick={() => sil(i.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><X size={14} /></button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
