import React from "react";
import { PackageCheck, Clock, Send } from "lucide-react";
import { projeTamamlanmaDurumu, sureFormatla, saatSayisi, guvenliCagir } from "./yardimcilar";

// Bu dosya bağımsız bir "iç modül"dür: Sipariş'in veya WBS'in içini bilmez,
// sadece tumWbs + malzemeTalepleri okuyup projelendirmenin ne durumda olduğunu
// gösterir ve tamamlandığında "planlama_paketler" koleksiyonuna MÜHÜRLENMİŞ bir
// paket yazar. Gelecekteki Çizelgeleme (Planlama/APS) modülü, WBS ağacının
// içine hiç girmeden sadece bu koleksiyonu okuyarak çalışabilir — bağımlılık
// tek yönlüdür: Sipariş -> Projelendirme(WBS+Rota+Malzeme) -> [bu modül] -> Çizelgeleme.
export default function ProjelendirmeOzeti({ api, ui, yazabilir, siparis, tumWbs, malzemeTalepleri, paketler }) {
  const s = ui.stil || {};
  const durum = projeTamamlanmaDurumu(siparis.id, tumWbs, malzemeTalepleri);
  const mevcutPaket = (paketler || []).find((p) => p.siparisId === siparis.id);
  const zatenGonderildi = siparis.durum === "Planlamada" || !!mevcutPaket;

  const simdi = new Date().toISOString();
  const gecenSure = siparis.projelendirmeBaslangic
    ? sureFormatla(siparis.projelendirmeBaslangic, mevcutPaket ? mevcutPaket.projelendirmeBitis : simdi)
    : null;

  async function planlamayaGonder() {
    if (!durum.hepsiTamam) return;
    const bitisIso = new Date().toISOString();
    const wbsOzet = (tumWbs || [])
      .filter((w) => w.siparisId === siparis.id)
      .map((w) => ({
        id: w.id, wbsKodu: w.wbsKodu, parentId: w.parentId,
        parcaNo: w.parcaNo, parcaAdi: w.parcaAdi, revizyon: w.revizyon,
        miktar: w.miktar, rotaId: w.rotaId, requiredMachineClass: w.requiredMachineClass,
        tightestTolerance: w.tightestTolerance,
      }));

    const tamam = await guvenliCagir(async () => {
      await api.ekleNumarali("planlama_paketler", `PKT-${siparis.id}`, {
        siparisId: siparis.id,
        musteriAdi: siparis.musteriAdi,
        projeKodu: siparis.projeKodu,
        toplamKalem: durum.toplam,
        wbsOzet,
        projelendirmeBaslangic: siparis.projelendirmeBaslangic || bitisIso,
        projelendirmeBitis: bitisIso,
        projelendirmeSuresiSaat: saatSayisi(siparis.projelendirmeBaslangic, bitisIso),
        teslimTarihi: siparis.teslimTarihi || null,
        durum: "PlanlamayaHazir",
      });
      await api.guncelle("planlama_siparisler", siparis.id, {
        durum: "Planlamada",
        projelendirmeBitis: bitisIso,
      });
    }, "Paket planlamaya gönderilemedi");
    return tamam;
  }

  return (
    <div className="card" style={{ padding: 16, marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
        <h3 style={{ margin: 0, color: s.renk?.yazi }}>Projelendirme İlerlemesi</h3>
        {gecenSure && (
          <span style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: s.renk?.soluk }}>
            <Clock size={14} /> {zatenGonderildi ? "Projelendirme süresi" : "Geçen süre"}: <b style={{ color: s.renk?.yazi }}>{gecenSure}</b>
          </span>
        )}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 10, marginBottom: 14 }}>
        <ui.Stat label="Toplam WBS Kalemi" value={String(durum.toplam)} />
        <ui.Stat label="Kırılım Tamam" value={`${durum.kirilimTamam}/${durum.toplam}`} highlight={durum.toplam > 0 && durum.kirilimTamam === durum.toplam} />
        <ui.Stat label="Rota Tanımlı" value={`${durum.rotaTamam}/${durum.toplam}`} highlight={durum.toplam > 0 && durum.rotaTamam === durum.toplam} />
        <ui.Stat label="Malzeme Kararı Verilmiş" value={`${durum.malzemeTamam}/${durum.toplam}`} highlight={durum.toplam > 0 && durum.malzemeTamam === durum.toplam} />
      </div>

      {zatenGonderildi ? (
        <div style={{ display: "flex", alignItems: "center", gap: 8, padding: 10, borderRadius: 8, background: (s.renk?.iyi || "#22c55e") + "1a", border: `1px solid ${(s.renk?.iyi || "#22c55e")}55`, color: s.renk?.iyi || "#22c55e" }}>
          <PackageCheck size={16} /> Bu sipariş paketlenip planlamaya devredildi. WBS artık sadece izlenebilir referans.
        </div>
      ) : durum.hepsiTamam ? (
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, padding: 10, borderRadius: 8, background: (s.renk?.iyi || "#22c55e") + "1a", border: `1px solid ${(s.renk?.iyi || "#22c55e")}55`, flexWrap: "wrap" }}>
          <span style={{ color: s.renk?.iyi || "#22c55e", fontWeight: 600 }}>
            <PackageCheck size={16} style={{ verticalAlign: -3, marginRight: 6 }} /> Tüm kalemler tamam — planlamaya devredilebilir.
          </span>
          {yazabilir && (
            <button className="btn-ghost" style={s.anaDugme} onClick={planlamayaGonder}>
              <Send size={14} style={{ marginRight: 6 }} /> Planlamaya Gönder
            </button>
          )}
        </div>
      ) : (
        <div style={{ padding: 10, borderRadius: 8, background: (s.renk?.uyari || "#facc15") + "1a", border: `1px solid ${(s.renk?.uyari || "#facc15")}55`, color: s.renk?.uyari || "#facc15", fontSize: 13 }}>
          Henüz tamam değil — her WBS kaleminde kırılım, rota ve malzeme kararını tamamlayın.
        </div>
      )}
    </div>
  );
}
