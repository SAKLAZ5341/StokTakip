import React, { useEffect, useState } from "react";
import SiparisPaneli from "./SiparisPaneli";
import WbsPaneli from "./WbsPaneli";
import ProjelendirmeOzeti from "./ProjelendirmeOzeti";
import IsEmirleriPaneli from "./IsEmirleriPaneli";
import { kokWbsKoduUret, guvenliCagir } from "./yardimcilar";

// Giriş noktası. BENIOKU.md sözleşmesi gereği ana programın hiçbir dosyasını
// import etmez; ihtiyacı olan her şey (kullanici, yetki, api, ui, veri) prop
// olarak gelir.
export default function Planlama({ kullanici, yetki, api, ui, veri }) {
  const yazabilir = yetki === "duzenle";
  const s = ui.stil || {};

  const [siparisler, setSiparisler] = useState([]);
  const [tumWbs, setTumWbs] = useState([]);
  const [rotalar, setRotalar] = useState([]);
  const [malzemeTalepleri, setMalzemeTalepleri] = useState([]);
  const [fasonTalepleri, setFasonTalepleri] = useState([]);
  const [paketler, setPaketler] = useState([]);
  const [isEmirleri, setIsEmirleri] = useState([]);
  const [ayarlarListesi, setAyarlarListesi] = useState([]);
  const [seciliSiparisId, setSeciliSiparisId] = useState(null);

  useEffect(() => { const durdur = api.dinle("planlama_siparisler", setSiparisler); return durdur; }, []);
  useEffect(() => { const durdur = api.dinle("planlama_wbs", setTumWbs); return durdur; }, []);
  useEffect(() => { const durdur = api.dinle("planlama_rotalar", setRotalar); return durdur; }, []);
  useEffect(() => { const durdur = api.dinle("planlama_malzemeTalepleri", setMalzemeTalepleri); return durdur; }, []);
  useEffect(() => { const durdur = api.dinle("planlama_fasonTalepleri", setFasonTalepleri); return durdur; }, []);
  // planlama_paketler: Projelendirme tamamlanınca mühürlenen, Çizelgeleme
  // (gelecek modül) için tek doğruluk kaynağı olacak devir paketleri.
  useEffect(() => { const durdur = api.dinle("planlama_paketler", setPaketler); return durdur; }, []);
  // planlama_isEmirleri: AS9100 iş emirleri (SERİ/FAI/DELTA/FAI_REWORK/MONTAJ),
  // her biri kendi istasyon atamalarını taşır — tarihler burada SAKLANMAZ,
  // cizelgeMotoru.js her seferinde canlı hesaplar.
  useEffect(() => { const durdur = api.dinle("planlama_isEmirleri", setIsEmirleri); return durdur; }, []);
  // planlama_ayarlar: modülün kendi sahip olduğu basit takvim ayarı (günlük
  // kapasite saat) — veri prop'unda vardiya/takvim kaynağı olmadığı için.
  useEffect(() => { const durdur = api.dinle("planlama_ayarlar", setAyarlarListesi); return durdur; }, []);

  const seciliSiparis = siparisler.find((sp) => sp.id === seciliSiparisId) || null;

  async function projelendir(siparis) {
    if (siparis.durum === "Projelendirildi") return;
    const kokVarMi = tumWbs.some((w) => w.siparisId === siparis.id && !w.parentId);
    const tamam = await guvenliCagir(async () => {
      if (!kokVarMi) {
        await api.ekle("planlama_wbs", {
          siparisId: siparis.id,
          parentId: null,
          wbsKodu: kokWbsKoduUret(siparis),
          parcaNo: siparis.parcaNo,
          parcaAdi: siparis.parcaAdi,
          revizyon: siparis.revizyon || "00",
          miktar: siparis.miktar,
          tightestTolerance: siparis.tightestTolerance || "",
          requiredMachineClass: siparis.requiredMachineClass || "",
          malzemeGerekli: true,
          fasonGerekli: false,
          faiGerekli: !!siparis.sfc,
          durum: "Beklemede",
          faiOnay: false,
          faiDurum: "beklemede",
        });
      }
      await api.guncelle("planlama_siparisler", siparis.id, {
        durum: "Projelendirildi",
        projeDeadline: siparis.teslimTarihi || null,
        projelendirmeBaslangic: new Date().toISOString(),
      });
    }, "Sipariş projelendirilemedi");
    if (tamam) setSeciliSiparisId(siparis.id);
  }

  return (
    <div style={{ padding: 16 }}>
      <div style={{ marginBottom: 16 }}>
        <h2 style={{ margin: 0, color: s.renk?.yazi }}>Planlama — Açık Siparişler ve Proje (WBS) Yönetimi</h2>
        <p style={{ color: s.renk?.soluk, margin: "4px 0 0" }}>
          {kullanici?.email} · Yetki: {yazabilir ? "Düzenle" : "Görüntüle"}
        </p>
      </div>

      <SiparisPaneli
        api={api} ui={ui} veri={veri} yazabilir={yazabilir}
        siparisler={siparisler} seciliSiparisId={seciliSiparisId}
        onSec={setSeciliSiparisId} onProjelendir={projelendir}
      />

      {seciliSiparis && (seciliSiparis.durum === "Açık" || seciliSiparis.durum === "Projelendirildi") && (
        <>
          <ProjelendirmeOzeti
            api={api} ui={ui} yazabilir={yazabilir}
            siparis={seciliSiparis} tumWbs={tumWbs}
            malzemeTalepleri={malzemeTalepleri} paketler={paketler}
          />
          <WbsPaneli
            api={api} ui={ui} veri={veri} yazabilir={yazabilir}
            siparis={seciliSiparis} tumWbs={tumWbs} rotalar={rotalar}
            malzemeTalepleri={malzemeTalepleri} fasonTalepleri={fasonTalepleri}
          />
        </>
      )}

      {seciliSiparis && seciliSiparis.durum === "Planlamada" && (
        <IsEmirleriPaneli
          api={api} ui={ui} veri={veri} yazabilir={yazabilir}
          siparis={seciliSiparis} tumSiparisler={siparisler} tumWbs={tumWbs}
          tumIsEmirleri={isEmirleri} malzemeTalepleri={malzemeTalepleri} rotalar={rotalar}
          ayarlarListesi={ayarlarListesi}
        />
      )}

      {!seciliSiparis && (
        <div className="card" style={{ padding: 24, textAlign: "center", color: s.renk?.soluk }}>
          WBS / rota / çizelge detaylarını görmek için yukarıdan bir sipariş seçin.
        </div>
      )}
    </div>
  );
}
