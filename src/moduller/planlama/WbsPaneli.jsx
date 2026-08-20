import React, { useMemo, useState } from "react";
import { Plus, ChevronDown, ChevronRight, GitBranch, PackageSearch, ClipboardCheck, Layers, Truck, Factory } from "lucide-react";
import { hazirlikHesapla, altWbsKoduUret, kokWbsKoduUret, guvenliCagir } from "./yardimcilar";
import RotaPaneli from "./RotaPaneli";

const RENK_ANAHTAR = { yesil: "iyi", sari: "uyari", kirmizi: "hata", gri: "soluk" };
const RENK_YEDEK = { yesil: "#15803d", sari: "#a16207", kirmizi: "#dc2626", gri: "#64748b" };

function rozetRengi(ui, renkAnahtari) {
  const s = ui.stil || {};
  return (s.renk && s.renk[RENK_ANAHTAR[renkAnahtari]]) || RENK_YEDEK[renkAnahtari];
}

function Rozet({ harf, renkAnahtari, ui, baslik }) {
  const renk = rozetRengi(ui, renkAnahtari);
  return (
    <span
      title={baslik}
      style={{
        display: "inline-flex", alignItems: "center", justifyContent: "center",
        width: 20, height: 20, borderRadius: "50%", fontSize: 11, fontWeight: 700,
        background: renk + "26", color: renk, border: `1px solid ${renk}66`, marginRight: 4,
      }}
    >
      {harf}
    </span>
  );
}

// WBS ağacını (sipariş bazlı) hiyerarşik listeler; her satırda M/S/F/Q rozetleri,
// alt parça / rota / malzeme talebi / FAI onay eylemleri bulunur.
export default function WbsPaneli({ api, ui, veri, yazabilir, siparis, tumWbs, rotalar, malzemeTalepleri, fasonTalepleri }) {
  const [acikNodlar, setAcikNodlar] = useState(() => new Set());
  const [ekleFormu, setEkleFormu] = useState(null); // { parentId, parentKodu } | null
  const [rotaModal, setRotaModal] = useState(null); // node | null
  const [malzemeModal, setMalzemeModal] = useState(null); // node | null
  const s = ui.stil || {};

  const cocuklariAl = (id) => (tumWbs || []).filter((w) => w.parentId === id);
  const kokler = useMemo(
    () => (tumWbs || []).filter((w) => w.siparisId === siparis.id && !w.parentId),
    [tumWbs, siparis.id]
  );

  function acKapa(id) {
    setAcikNodlar((prev) => {
      const yeni = new Set(prev);
      if (yeni.has(id)) yeni.delete(id); else yeni.add(id);
      return yeni;
    });
  }

  async function faiOnayla(node, deger) {
    await guvenliCagir(
      () => api.guncelle("planlama_wbs", node.id, { faiOnay: deger, faiDurum: deger ? "onaylandi" : "beklemede" }),
      "FAI durumu güncellenemedi"
    );
  }

  // Tek tık: "bu kalemin (ve varsa alt kırılımının) tanımı bitti" onayı.
  // Projelendirme tamamlanma kapısının 3 şartından biri budur.
  async function kirilimTamamToggle(node) {
    await guvenliCagir(
      () => api.guncelle("planlama_wbs", node.id, { bomTamamlandi: !node.bomTamamlandi }),
      "Kırılım durumu güncellenemedi"
    );
  }

  // Tek tık: hammadde standart satınalma ile mi geliyor, yoksa müşteri
  // (free-issue) mi gönderiyor. "Müşteri" seçilirse malzeme talebine gerek kalmaz.
  async function malzemeKaynagiToggle(node) {
    const yeni = node.malzemeKaynagi === "musteri" ? "standart" : "musteri";
    await guvenliCagir(
      () => api.guncelle("planlama_wbs", node.id, { malzemeKaynagi: yeni }),
      "Malzeme kaynağı güncellenemedi"
    );
  }

  function Satir({ node, derinlik }) {
    const altlar = cocuklariAl(node.id);
    const acik = acikNodlar.has(node.id);
    const hazirlik = hazirlikHesapla(node, tumWbs, malzemeTalepleri, fasonTalepleri);
    return (
      <>
        <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 4px", borderBottom: `1px solid ${s.renk?.kenar || "#d5dfec"}` }}>
          <div style={{ width: derinlik * 20, flexShrink: 0 }} />
          <button
            onClick={() => altlar.length && acKapa(node.id)}
            style={{ background: "none", border: "none", cursor: altlar.length ? "pointer" : "default", color: s.renk?.soluk, padding: 0 }}
          >
            {altlar.length > 0 ? (acik ? <ChevronDown size={16} /> : <ChevronRight size={16} />) : <span style={{ display: "inline-block", width: 16 }} />}
          </button>
          <span style={{ fontFamily: "monospace", fontSize: 12, color: s.renk?.vurgu, minWidth: 100 }}>{node.wbsKodu}</span>
          <span style={{ flex: 1, minWidth: 160 }}>
            {node.parcaAdi}{" "}
            <span style={{ color: s.renk?.soluk, fontSize: 12 }}>
              ({node.parcaNo} · Rev {node.revizyon || "00"})
            </span>
          </span>
          <span style={{ fontSize: 12, color: s.renk?.soluk, width: 60, textAlign: "right" }}>{node.miktar} ad.</span>
          <div style={{ display: "flex", alignItems: "center" }}>
            <Rozet harf="M" renkAnahtari={hazirlik.malzeme} ui={ui} baslik="Malzeme" />
            <Rozet harf="S" renkAnahtari={hazirlik.altParca} ui={ui} baslik="Alt Parça" />
            <Rozet harf="F" renkAnahtari={hazirlik.fason} ui={ui} baslik="Fason" />
            <Rozet harf="Q" renkAnahtari={hazirlik.kalite} ui={ui} baslik="Kalite / FAI" />
          </div>
          <div style={{ display: "flex", gap: 4 }}>
            {yazabilir && (
              <button
                className="btn-ghost" style={{ ...s.dugme, padding: "4px 8px" }}
                onClick={() => setEkleFormu({ parentId: node.id, parentKodu: node.wbsKodu })}
                title="Alt parça ekle"
              >
                <Plus size={14} />
              </button>
            )}
            <button className="btn-ghost" style={{ ...s.dugme, padding: "4px 8px" }} onClick={() => setRotaModal(node)} title="Rota / İstasyon">
              <GitBranch size={14} />
            </button>
            {yazabilir && node.malzemeGerekli && (
              <>
                <button
                  className="btn-ghost" style={{ ...s.dugme, padding: "4px 8px" }}
                  onClick={() => setMalzemeModal(node)} title="Malzeme talebi gönder"
                >
                  <PackageSearch size={14} />
                </button>
                <button
                  className="btn-ghost"
                  style={{ ...s.dugme, padding: "4px 8px", color: node.malzemeKaynagi === "musteri" ? (s.renk?.vurgu) : s.renk?.soluk }}
                  onClick={() => malzemeKaynagiToggle(node)}
                  title={node.malzemeKaynagi === "musteri" ? "Malzeme: Müşteriden geliyor — Standarta çevirmek için tıkla" : "Malzeme: Standart satınalma — Müşteri malzemesi yapmak için tıkla"}
                >
                  {node.malzemeKaynagi === "musteri" ? <Truck size={14} /> : <Factory size={14} />}
                </button>
              </>
            )}
            {yazabilir && (
              <button
                className="btn-ghost"
                style={{ ...s.dugme, padding: "4px 8px", color: node.bomTamamlandi ? (s.renk?.iyi || "#15803d") : s.renk?.soluk }}
                onClick={() => kirilimTamamToggle(node)}
                title={node.bomTamamlandi ? "Kırılım tamam işaretli — geri almak için tıkla" : "Bu kalemin (ve alt kırılımının) tanımı bitti mi? Tamamsa tıkla"}
              >
                <Layers size={14} />
              </button>
            )}
            {yazabilir && node.faiGerekli && (
              <button
                className="btn-ghost"
                style={{ ...s.dugme, padding: "4px 8px", color: node.faiOnay ? (s.renk?.iyi || "#15803d") : (s.renk?.uyari || "#a16207") }}
                onClick={() => faiOnayla(node, !node.faiOnay)}
                title={node.faiOnay ? "FAI onaylı — kaldırmak için tıkla" : "FAI onayla"}
              >
                <ClipboardCheck size={14} />
              </button>
            )}
          </div>
        </div>
        {acik && altlar.map((c) => <Satir key={c.id} node={c} derinlik={derinlik + 1} />)}
      </>
    );
  }

  return (
    <div className="card" style={{ padding: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, flexWrap: "wrap", gap: 8 }}>
        <h3 style={{ margin: 0, color: s.renk?.yazi }}>WBS / Ürün Ağacı — {siparis.musteriAdi}</h3>
        {yazabilir && (
          <button className="btn-ghost" style={s.anaDugme} onClick={() => setEkleFormu({ parentId: null, parentKodu: null })}>
            <Plus size={16} style={{ marginRight: 6 }} /> Ana Kalem Ekle
          </button>
        )}
      </div>
      <div style={{ fontSize: 12, color: s.renk?.soluk, marginBottom: 12 }}>
        Rozetler: <b>M</b> Malzeme · <b>S</b> Alt Parça · <b>F</b> Fason · <b>Q</b> Kalite/FAI —
        yeşil: hazır/geçerli değil, sarı: sürüyor, kırmızı: eksik, gri: bu kalemde geçerli değil.
        {" · "}<Layers size={12} style={{ verticalAlign: -2 }} /> Kırılım Tamam
        {" · "}<Factory size={12} style={{ verticalAlign: -2 }} />/<Truck size={12} style={{ verticalAlign: -2 }} /> Malzeme: Standart / Müşteri
      </div>

      <div style={{ overflowX: "auto" }}>
        {kokler.length === 0 && (
          <div style={{ color: s.renk?.soluk, padding: 12 }}>
            Bu sipariş için henüz WBS kalemi yok. "Ana Kalem Ekle" ile başlayın.
          </div>
        )}
        {kokler.map((k) => <Satir key={k.id} node={k} derinlik={0} />)}
      </div>

      {ekleFormu && (
        <WbsEkleFormu
          api={api} ui={ui} siparis={siparis}
          parentId={ekleFormu.parentId} parentKodu={ekleFormu.parentKodu}
          kardesSayisi={cocuklariAl(ekleFormu.parentId).length}
          kapat={() => setEkleFormu(null)}
        />
      )}

      {rotaModal && (
        <ui.Pencere
          acik baslik={`Rota / İstasyon — ${rotaModal.wbsKodu}`} genislik="820px"
          kapat={() => setRotaModal(null)}
          butonlar={[{ etiket: "Kapat", onTikla: () => setRotaModal(null) }]}
        >
          <RotaPaneli api={api} ui={ui} veri={veri} yazabilir={yazabilir} node={rotaModal} siparis={siparis} rotalar={rotalar} />
        </ui.Pencere>
      )}

      {malzemeModal && (
        <MalzemeTalepFormu api={api} ui={ui} veri={veri} node={malzemeModal} kapat={() => setMalzemeModal(null)} />
      )}
    </div>
  );
}

function WbsEkleFormu({ api, ui, siparis, parentId, parentKodu, kardesSayisi, kapat }) {
  const kokMu = !parentId;
  const [form, setForm] = useState({
    parcaNo: kokMu ? siparis.parcaNo : "",
    parcaAdi: kokMu ? siparis.parcaAdi : "",
    revizyon: kokMu ? (siparis.revizyon || "00") : "00",
    miktar: kokMu ? siparis.miktar : 1,
    tightestTolerance: kokMu ? (siparis.tightestTolerance || "") : "",
    requiredMachineClass: kokMu ? (siparis.requiredMachineClass || "") : "",
    malzemeGerekli: true,
    fasonGerekli: false,
    faiGerekli: kokMu ? siparis.sfc === true : false,
  });
  const s = ui.stil || {};

  function alan(k, v) { setForm((f) => ({ ...f, [k]: v })); }

  async function kaydet() {
    if (!form.parcaNo || !form.parcaAdi || !form.miktar) {
      alert("Parça No, Parça Adı ve Miktar zorunludur.");
      return;
    }
    const wbsKodu = parentId ? altWbsKoduUret(parentKodu, kardesSayisi) : kokWbsKoduUret(siparis);
    const tamam = await guvenliCagir(
      () =>
        api.ekle("planlama_wbs", {
          siparisId: siparis.id,
          parentId: parentId || null,
          wbsKodu,
          ...form,
          miktar: Number(form.miktar),
          durum: "Beklemede",
          faiOnay: false,
          faiDurum: "beklemede",
          malzemeKaynagi: "standart",
          bomTamamlandi: false,
        }),
      "WBS kalemi kaydedilemedi"
    );
    if (tamam) kapat();
  }

  return (
    <ui.Pencere
      acik
      baslik={parentId ? `Alt Parça Ekle (${parentKodu} altına)` : "Ana Kalem Ekle"}
      genislik="560px"
      kapat={kapat}
      butonlar={[{ etiket: "Vazgeç", onTikla: kapat }, { etiket: "Kaydet", onTikla: kaydet, birincil: true }]}
    >
      <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 10 }}>
        <label className="field-label">Parça No
          <input className="input" style={s.giris} value={form.parcaNo} onChange={(e) => alan("parcaNo", e.target.value)} />
        </label>
        <label className="field-label">Parça Adı
          <input className="input" style={s.giris} value={form.parcaAdi} onChange={(e) => alan("parcaAdi", e.target.value)} />
        </label>
        <label className="field-label">Revizyon
          <input className="input" style={s.giris} value={form.revizyon} onChange={(e) => alan("revizyon", e.target.value)} />
        </label>
        <label className="field-label">Miktar
          <input className="input" type="number" min="1" style={s.giris} value={form.miktar} onChange={(e) => alan("miktar", e.target.value)} />
        </label>
        <label className="field-label">Dar Tolerans (mm)
          <input className="input" style={s.giris} placeholder="ör. 0.01" value={form.tightestTolerance} onChange={(e) => alan("tightestTolerance", e.target.value)} />
        </label>
        <label className="field-label">Gerekli Makine Sınıfı
          <input className="input" style={s.giris} placeholder="ör. CNC3 / CNCT / TASLAMA" value={form.requiredMachineClass} onChange={(e) => alan("requiredMachineClass", e.target.value)} />
        </label>
        <label className="field-label" style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <input type="checkbox" checked={form.malzemeGerekli} onChange={(e) => alan("malzemeGerekli", e.target.checked)} /> Hammadde/malzeme gerekli
        </label>
        <label className="field-label" style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <input type="checkbox" checked={form.fasonGerekli} onChange={(e) => alan("fasonGerekli", e.target.checked)} /> Fason işlem gerekli
        </label>
        <label className="field-label" style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <input type="checkbox" checked={form.faiGerekli} onChange={(e) => alan("faiGerekli", e.target.checked)} /> FAI / AS9102 gerekli
        </label>
      </div>
    </ui.Pencere>
  );
}

function MalzemeTalepFormu({ api, ui, veri, node, kapat }) {
  const [stokKodu, setStokKodu] = useState("");
  const [stokAdi, setStokAdi] = useState("");
  const [miktar, setMiktar] = useState(node.miktar || 1);
  const [birim, setBirim] = useState("adet");
  const [secimAcik, setSecimAcik] = useState(false);
  const s = ui.stil || {};

  async function gonder() {
    if (!stokKodu || !miktar) { alert("Stok kodu ve miktar zorunludur."); return; }
    const tamam = await guvenliCagir(
      () =>
        api.ekle("planlama_malzemeTalepleri", {
          wbsId: node.id, siparisId: node.siparisId,
          stokKodu, stokAdi, miktar: Number(miktar), birim,
          durum: "beklemede", talepTarihi: ui.tarih.bugun(),
        }),
      "Malzeme talebi gönderilemedi"
    );
    if (tamam) kapat();
  }

  return (
    <ui.Pencere
      acik baslik={`Malzeme Talebi — ${node.wbsKodu}`} genislik="480px" kapat={kapat}
      butonlar={[{ etiket: "Vazgeç", onTikla: kapat }, { etiket: "Satınalmaya Gönder", onTikla: gonder, birincil: true }]}
    >
      <p style={{ color: s.renk?.soluk, fontSize: 13, marginTop: 0 }}>
        Bu talep <code>planlama_malzemeTalepleri</code> koleksiyonuna düşer. Gerçek satınalma sürecine
        (satinalma_talepleri) aktarımı, modül sözleşmesi gereği ana programda ayrıca bağlanmalıdır.
      </p>
      <div style={{ display: "grid", gap: 10 }}>
        <label className="field-label">Stok Kodu
          <input className="input" style={s.giris} value={stokKodu} onChange={(e) => setStokKodu(e.target.value)} />
        </label>
        <label className="field-label">Stok Adı
          <input className="input" style={s.giris} value={stokAdi} onChange={(e) => setStokAdi(e.target.value)} />
        </label>
        <div style={{ display: "flex", gap: 10 }}>
          <label className="field-label" style={{ flex: 1 }}>Miktar
            <input className="input" type="number" style={s.giris} value={miktar} onChange={(e) => setMiktar(e.target.value)} />
          </label>
          <label className="field-label" style={{ flex: 1 }}>Birim
            <input className="input" style={s.giris} value={birim} onChange={(e) => setBirim(e.target.value)} />
          </label>
        </div>
        {veri.stokKartlari && veri.stokKartlari.length > 0 && (
          <button type="button" className="btn-ghost" style={s.dugme} onClick={() => setSecimAcik((v) => !v)}>
            Stok Kartlarından Seç
          </button>
        )}
      </div>
      {secimAcik && (
        <div style={{ marginTop: 10, maxHeight: 220, overflowY: "auto", border: `1px solid ${s.renk?.kenar || "#d5dfec"}`, borderRadius: 8 }}>
          {(veri.stokKartlari || []).map((sk) => (
            <div
              key={sk.id || sk.kod}
              onClick={() => { setStokKodu(sk.kod || sk.stokKodu || ""); setStokAdi(sk.ad || sk.stokAdi || ""); setBirim(sk.birim || "adet"); setSecimAcik(false); }}
              style={{ padding: 8, cursor: "pointer", borderBottom: `1px solid ${s.renk?.kenar || "#d5dfec"}` }}
            >
              <strong>{sk.kod || sk.stokKodu}</strong> — {sk.ad || sk.stokAdi}
            </div>
          ))}
        </div>
      )}
    </ui.Pencere>
  );
}
